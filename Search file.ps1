#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Search for files with a matching Filter in the file name.

    .DESCRIPTION
        This script will report the quantity of files that contain a specific
        key Filter in their file name.

    .PARAMETER Tasks
        Collection of a items to check.

    .PARAMETER Tasks.ComputerName
        Where the search is executed. If left blank it's best to UNC paths in
        'FolderPath'.

    .PARAMETER Tasks.FolderPath
        One or more paths to a folder where to search for files matching the 
        filter.

    .PARAMETER Tasks.Filter
        The filter used to find matching files. This filter is directly passed 
        to 'Get-ChildItem -Filter'.

        Examples:
        - '*.ps1'  : find all files with extension '.ps1'
        - '*kiwi*' : find all files with the word 'kiwi' in the file name

    .PARAMETER Recurse
        When the true the parent and child folders are searched for matching
        files. When false only the parent folder is searched.

    .PARAMETER Tasks.SendMail.To
        List of e-mail addresses where to send the e-mail too.

    .PARAMETER Tasks.SendMail.When
        When to send an e-mail. 
        
        Valid options:
        - Always                : Always send an e-mail, even without matches
        - OnlyWhenFilesAreFound : Only send an e-mail when matches are found
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [string]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Search file\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)
        
Begin {
    Function Get-JobDurationHC {
        [OutputType([TimeSpan])]
        Param (
            [Parameter(Mandatory)]
            [System.Management.Automation.Job]$Job
        )

        $params = @{
            Start = $Job.PSBeginTime
            End   = $Job.PSEndTime
        }
        $jobDuration = New-TimeSpan @params

        #region Get computer name
        $computerName = $Job.Location
        if ($Job.Location -eq 'localhost') {
            $computerName = $env:COMPUTERNAME
        }
        #endregion

        $M = "'{0}' job duration '{1:hh}:{1:mm}:{1:ss}:{1:fff}'" -f 
        $computerName, $jobDuration
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $jobDuration
    }
    Function Get-JobResultsAndErrorsHC {
        [OutputType([PSCustomObject])]
        Param (
            [Parameter(Mandatory)]
            [System.Management.Automation.Job]$Job
        )

        $result = [PSCustomObject]@{
            Result = $null
            Errors = @()
        }

        #region Get computer name
        $computerName = $Job.Location
        if ($Job.Location -eq 'localhost') {
            $computerName = $env:COMPUTERNAME
        }
        #endregion

        #region Get job results
        $M = "'{0}' job '{1}'" -f $computerName, $job.State
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
              
        $jobErrors = @()
        $receiveParams = @{
            ErrorVariable = 'jobErrors'
            ErrorAction   = 'SilentlyContinue'
        }
        $result.Result = $Job | Receive-Job @receiveParams
        #endregion
   
        #region Get job errors
        foreach ($e in $jobErrors) {
            $M = "'{0}' job error '{1}'" -f $computerName, $e.ToString()
            Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
                  
            $result.Errors += $e.ToString()
            $error.Remove($e)
        }
        if ($resultErrors = $result.Result.Error | Where-Object { $_ }) {
            foreach ($e in $resultErrors) {
                $M = "'{0}' error '{1}'" -f $computerName, $e
                Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
                
                $result.Errors += $e
            }
        }
        #endregion

        $result.Result = $result.Result | 
        Select-Object -Property * -ExcludeProperty 'Error'

        if ((-not $result.Errors) -and (-not $result.Result.Error)) {
            $M = "'{0}' job successful found '{1}' files" -f $computerName,
            $result.Result.Files.Count
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        }

        $result
    }

    $getMatchingFilesHC = {
        [OutputType([PSCustomObject])]
        Param (
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            [Boolean]$Recurse,
            [Parameter(Mandatory)]
            [String[]]$Filters
        )

        if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
            throw "Path '$Path' not found."
        }
        
        foreach ($filter in $Filters) {
            try {
                $startDate = Get-Date

                $result = [PSCustomObject]@{
                    Filter   = $filter
                    Files    = @()
                    Duration = $null
                    Error    = $null
                }
                
                $params = @{
                    LiteralPath = $Path 
                    Recurse     = $Recurse
                    Filter      = $filter
                    File        = $true
                    ErrorAction = 'Stop'
                }
                $result.Files += Get-ChildItem @params
            }
            catch {
                $result.Error = "Failed retrieving files with filter '$filter': $_"
                $Error.RemoveAt(0)
            }
            finally {
                $result.Duration = (Get-Date) - $startDate
                $result
            }
        }
    }
    $getJobResult = {
        #region Verbose
        $M = "'{0}' Get job result for Path '{1}'" -f 
        $completedJob.ComputerName, $completedJob.Path
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region Get job results
        $params = @{
            Job = $completedJob.Job.Object
        }
        $jobOutput = Get-JobResultsAndErrorsHC @params

        $completedJob.Job.Duration = Get-JobDurationHC @params 
        #endregion

        #region Add job results
        $completedJob.Job.Result = $jobOutput.Result

        $jobOutput.Errors | ForEach-Object { 
            $completedJob.Job.Errors += $_ 
        }
        #endregion
            
        $completedJob.Job.Object = $null
    }

    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start
    
        $error.Clear()
        Get-Job | Remove-Job -Force -EA Ignore
        
        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
        
        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion
        
        #region Test .json file properties
        if (-not ($Tasks = $file.Tasks)) {
            throw "Input file '$ImportFile': Property 'Tasks' not found."
        }

        foreach ($task in $Tasks) {
            if (-not $task.FolderPath) {
                throw "Input file '$ImportFile': Property 'FolderPath' is mandatory."
            }
            if (-not $task.Filter) {
                throw "Input file '$ImportFile': Property 'Filter' is mandatory."
            }

            $task.ComputerName | Group-Object | 
            Where-Object { $_.Count -ge 2 } |
            ForEach-Object {
                throw "Input file '$ImportFile': Property 'ComputerName' contains the duplicate value '$($_.Name)'. Duplicate values are not allowed."
            }

            $task.Filter | Group-Object | Where-Object { $_.Count -ge 2 } |
            ForEach-Object {
                throw "Input file '$ImportFile': Property 'Filter' contains the duplicate value '$($_.Name)'. Duplicate values are not allowed."
            }

            $task.FolderPath | Group-Object | Where-Object { $_.Count -ge 2 } |
            ForEach-Object {
                throw "Input file '$ImportFile': Property 'FolderPath' contains the duplicate value '$($_.Name)'. Duplicate values are not allowed."
            }

            if ($task.PSObject.Properties.Name -notContains 'Recurse') {
                throw "Input file '$ImportFile': Property 'Recurse' is mandatory."
            }
            if (-not ($task.Recurse -is [boolean])) {
                throw "Input file '$ImportFile': The value '$($task.Recurse)' in 'Recurse' is not a true false value."
            }
            if (-not $task.SendMail) {
                throw "Input file '$ImportFile': Property 'SendMail' is mandatory."
            }
            if (-not $task.SendMail.To) {
                throw "Input file '$ImportFile': Property 'SendMail.To' is mandatory."
            }
            if (
                $task.SendMail.When -notMatch '^Always$|^OnlyWhenFilesAreFound$'
            ) {
                throw "Input file '$ImportFile': The value '$($task.SendMail.When)' in 'SendMail.When' is not supported. Only the value 'Always' or 'OnlyWhenFilesAreFound' can be used."
            }
        }

        if ($file.PSObject.Properties.Name -notContains 'MaxConcurrentJobs') {
            throw "Input file '$ImportFile': Property 'MaxConcurrentJobs' not found."
        }
        if (-not ($file.MaxConcurrentJobs -is [int])) {
            throw "Input file '$ImportFile': Property 'MaxConcurrentJobs' needs to be a number, the value '$($file.MaxConcurrentJobs)' is not supported."
        }

        $maxConcurrentJobs = [int]$file.MaxConcurrentJobs
        #endregion

        #region Add properties
        foreach ($task in $Tasks) {
            if (
                (-not $task.ComputerName) -or 
                ($task.ComputerName -eq 'localhost') -or
                ($task.ComputerName -eq "$env:COMPUTERNAME.$env:USERDNSDOMAIN")
            ) {
                $task.ComputerName = $env:COMPUTERNAME
            }

            $task | Add-Member -NotePropertyMembers @{
                Jobs = foreach ($computerName in $task.ComputerName) {
                    foreach ($path in $task.FolderPath) {
                        [PSCustomObject]@{
                            ComputerName = $computerName
                            Path         = $path
                            Job          = @{
                                Object   = $null
                                Duration = $null
                                Result   = $null
                                Errors   = @()
                            }
                        }
                    }
                }
            }
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}       
Process {
    Try {
        #region Start jobs
        foreach ($task in $Tasks) {
            foreach ($j in $task.Jobs) {
                #region Get matching files
                $invokeParams = @{
                    ScriptBlock  = $getMatchingFilesHC
                    ArgumentList = $j.Path, $task.Recurse, $task.Filter
                }
        
                $M = "'{0}' Start job for Path '{1}' Filter '{3}' Recurse '{2}'" -f 
                $j.ComputerName,
                $invokeParams.ArgumentList[0], 
                $invokeParams.ArgumentList[1],
                $($invokeParams.ArgumentList[2] -join ', ')
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        
                $j.Job.Object = if ($j.ComputerName -ne $env:COMPUTERNAME) {
                    $invokeParams.ComputerName = $j.ComputerName
                    $invokeParams.AsJob = $true
                    Invoke-Command @invokeParams
                }
                else {
                    Start-Job @invokeParams
                }
                
                $M = "'{0}' job '{1}'" -f $j.ComputerName, $j.Job.Object.State
                Write-Verbose $M
                #endregion

                #region Wait for max running jobs
                $waitParams = @{
                    Name       = $Tasks.Jobs.Job.Object | Where-Object { $_ }
                    MaxThreads = $maxConcurrentJobs
                }
                Wait-MaxRunningJobsHC @waitParams
                #endregion

                #region Get job results
                foreach (
                    $completedJob in 
                    $Tasks.Jobs | Where-Object {
                            ($_.Job.Object.State -match 'Completed|Failed')
                    }
                ) {
                    & $getJobResult
                }
                #endregion
            }
        }
        #endregion

        #region Wait for jobs to finish and get results
        while (
            $runningJobs = $Tasks.Jobs.Job.Object | Where-Object { $_ }
        ) {
            #region Verbose progress
            $runningJobCounter = ($runningJobs | Measure-Object).Count
            if ($runningJobCounter -eq 1) {
                $M = 'Wait for the last running job to finish'
            }
            else {
                $M = "Wait for one of '{0}' running jobs to finish" -f $runningJobCounter
            }
            Write-Verbose $M
            #endregion

            $finishedJob = $runningJobs | Wait-Job -Any

            $completedJob = $Tasks.Jobs | Where-Object {
                ($_.Job.Object.Id -eq $finishedJob.Id)
            }
            
            & $getJobResult
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}
End {
    try {
        for ($i = 0; $i -lt $Tasks.Count; $i++) {
            #region Verbose
            $M = "Task ComputerName '{0}' Path '{1}' Filter '{2}' Recurse '{3}' MailTo '{4}' MailWhen '{5}'" -f 
            $($Tasks[$i].ComputerName -join ', '), 
            $($Tasks[$i].FolderPath -join ', '), 
            $($Tasks[$i].Filter -join ', '), 
            $Tasks[$i].Recurse,
            $($Tasks[$i].SendMail.To -join ', '), 
            $Tasks[$i].SendMail.When
            Write-Verbose $M
            #endregion
        
            $mailParams = @{
                To        = $Tasks[$i].SendMail.To
                Bcc       = $ScriptAdmin
                Priority  = 'Normal'
                LogFolder = $logParams.LogFolder
                Header    = if ($Tasks[$i].SendMail.Header) {
                    $Tasks[$i].SendMail.Header
                }
                else { $ScriptName }
                Save      = "$logFile - $i - Mail.html"
            }

            $excelParams = @{
                Path         = "$logFile - $i - Log.xlsx"
                AutoSize     = $true
                FreezeTopRow = $true
            }
            $excelSheet = @{
                Files  = @()
                Errors = @()
            }

            #region Create Excel worksheet Files
            $excelSheet.Files += foreach ($j in $Tasks[$i].Jobs) {
                foreach ($r in $j.Job.Result) {
                    $r.Files | Select-Object -Property @{
                        Name       = 'ComputerName';
                        Expression = { $j.ComputerName }
                    },
                    @{
                        Name       = 'Path';
                        Expression = { $j.Path }
                    },
                    @{
                        Name       = 'Recurse';
                        Expression = { $Tasks[$i].Recurse }
                    },
                    @{
                        Name       = 'Filter';
                        Expression = { $r.Filter }
                    },
                    @{
                        Name       = 'File';
                        Expression = { $_.FullName }
                    },
                    @{
                        Name       = 'CreationTime';
                        Expression = { $_.CreationTime }
                    },
                    @{
                        Name       = 'LastWriteTime';
                        Expression = { $_.LastWriteTime }
                    },
                    @{
                        Name       = 'Size'; 
                        Expression = { [MATH]::Round($_.Length / 1GB, 2) } 
                    },
                    @{
                        Name       = 'Size_'; 
                        Expression = { $_.Length } 
                    },
                    @{
                        Name       = 'Duration';
                        Expression = { 
                            '{0:hh}:{0:mm}:{0:ss}:{0:fff}' -f $r.Duration 
                        }
                    }
                }
            }

            if ($excelSheet.Files) {
                $excelParams.WorksheetName = 'Files'
                $excelParams.TableName = 'Files'

                $M = "Export {0} rows to sheet '{1}' in Excel file '{2}'" -f 
                $excelSheet.Files.Count, 
                $excelParams.WorksheetName, $excelParams.Path
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                $excelSheet.Files | 
                Export-Excel @excelParams -AutoNameRange -CellStyleSB {
                    Param (
                        $WorkSheet,
                        $TotalRows,
                        $LastColumn
                    )
    
                    @($WorkSheet.Names['Size'].Style).ForEach( {
                            $_.NumberFormat.Format = '?\ \G\B'
                            $_.HorizontalAlignment = 'Center'
                        })
    
                    @($WorkSheet.Names['Size_'].Style).ForEach( {
                            $_.NumberFormat.Format = '?\ \B'
                            $_.HorizontalAlignment = 'Center'
                        })
                }

                $mailParams.Attachments = $excelParams.Path
            }
            #endregion

            #region Create Excel worksheet Errors
            $excelSheet.Errors += foreach ($j in $Tasks[$i].Jobs) {
                $j.Job | Where-Object { $_.Errors } | Select-Object -Property @{
                    Name       = 'ComputerName';
                    Expression = { $j.ComputerName }
                },
                @{
                    Name       = 'Path';
                    Expression = { $j.Path }
                },
                @{
                    Name       = 'Filter'; 
                    Expression = { $Tasks[$i].Filter -join ', ' }
                },
                @{
                    Name       = 'Duration';
                    Expression = { 
                        '{0:hh}:{0:mm}:{0:ss}:{0:fff}' -f $_.Duration 
                    }
                },
                @{
                    Name       = 'Error';
                    Expression = { $_.Errors -join ', ' }
                }
            }

            if ($excelSheet.Errors) {
                $excelParams.WorksheetName = 'Errors'
                $excelParams.TableName = 'Errors'

                $M = "Export {0} rows to sheet '{1}' in Excel file '{2}'" -f
                $excelSheet.Errors.Count, 
                $excelParams.WorksheetName, $excelParams.Path
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                $excelSheet.Errors | Export-Excel @excelParams

                $mailParams.Attachments = $excelParams.Path
            }
            #endregion

            #region Send mail
            if (
                ($Tasks[$i].SendMail.When -eq 'Always') -or
                ($excelSheet.Files) -or
                ($excelSheet.Errors)
            ) {
                $errorMessage = $null

                #region Subject and Priority
                $mailParams.Subject = '{0} file{1} found' -f 
                $excelSheet.Files.Count,
                $(if ($excelSheet.Files.Count -ne 1) { 's' })
                
                if ($excelSheet.Errors) {
                    $mailParams.Priority = 'High'

                    $mailParams.Subject += ', {0} error{1}' -f 
                    $excelSheet.Errors.Count,
                    $(if ($excelSheet.Errors.Count -ne 1) { 's' })

                    $errorMessage = "<p>Detected <b>{0} error{1}</b> during execution.</p>" -f 
                    $excelSheet.Errors.Count,
                    $(if ($excelSheet.Errors.Count -ne 1) { 's' })
                }
                #endregion


                $tableRows = foreach (
                    $computerName in 
                    $Tasks[$i].ComputerName
                ) {
                    foreach ($path in $Tasks[$i].FolderPath) {
                        $computerPathHtml = "<tr>
                            <th>{0}</th>
                            <th>{1}</th>
                       </tr>
                       <tr>
                            <td>Filter</td>
                            <td>Files found</td>
                        </tr>" -f $computerName, $(
                            if ($path -match '^\\\\') {
                                '<a href="{0}">{0}</a>' -f $path
                            }
                            else {
                                $uncPath = $path -Replace '^.{2}', (
                                    '\\{0}\{1}$' -f $computerName, $path[0]
                                )
                                '<a href="{0}">{0}</a>' -f $uncPath
                            }
                        )

                        $matchesFoundHtml = foreach (
                            $filter in 
                            $Tasks[$i].Filter
                        ) {
                            $matchesCount = $excelSheet.Files | Where-Object {
                                ($_.ComputerName -eq $computerName) -and
                                ($_.Path -eq $path) -and
                                ($_.Filter -eq $filter) 
                            } | Measure-Object |
                            Select-Object -ExpandProperty Count

                            if (
                                ($Tasks[$i].SendMail.When -eq 'Always') -or
                                ($matchesCount)
                            ) {
                                "<tr>
                                    <td>{0}</td>
                                    <td>{1}</td>
                                </tr>" -f $filter, $matchesCount    
                            }
                        }

                        if ($matchesFoundHtml) {
                            $computerPathHtml
                            $matchesFoundHtml                            
                        }

                    }
                }

                $mailParams.Message = "
                $errorMessage
                <p>Found a total of <b>{0} files</b>:</p>
                <table>
                    $tableRows
                </table>
                {1}" -f $excelSheet.Files.Count, $(
                    if ($mailParams.Attachments) {
                        '<p><i>* Check the attachment for details</i></p>'
                    }
                )
                
                Write-Verbose 'Send mail'
                Write-Verbose $mailParams.Message
                
                Get-ScriptRuntimeHC -Stop
                Send-MailHC @mailParams
            }
            #endregion
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Get-Job | Remove-Job -Force
        Write-EventLog @EventEndParams
    }
}