#Requires -Version 7
#Requires -Modules Toolbox.HTML, Toolbox.EventLog, Toolbox.Remoting

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

    .PARAMETER PSSessionConfiguration
        The version of PowerShell on the remote endpoint as returned by
        Get-PSSessionConfiguration.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [string]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$SearchScript = "$PSScriptRoot\Get file.ps1",
    [String]$PSSessionConfiguration = 'PowerShell.7',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Search file\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

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

        #region Test script path exists
        try {
            $params = @{
                Path        = $SearchScript
                ErrorAction = 'Stop'
            }
            $searchScriptPath = (Get-Item @params).FullName
        }
        catch {
            throw "Search script with path '$SearchScript' not found"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        try {
            @(
                'MaxConcurrentJobs', 'Tasks'
            ).where(
                { -not $file.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            try {
                $null = [int]$MaxConcurrentJobs = $file.MaxConcurrentJobs
            }
            catch {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$($file.MaxConcurrentJobs)' is not supported."
            }

            $Tasks = $file.Tasks

            foreach ($task in $Tasks) {
                @(
                    'FolderPath', 'Filter', 'ComputerName', 'SendMail'
                ).where(
                    { -not $task.$_ }
                ).foreach(
                    { throw "Property 'Tasks.$_' not found" }
                )

                if (-not ($task.Recurse -is [boolean])) {
                    throw "The value '$($task.Recurse)' in 'Recurse' is not a true false value."
                }

                @(
                    'To', 'When'
                ).where(
                    { -not $task.SendMail.$_ }
                ).foreach(
                    { throw "Property 'Tasks.SendMail.$_' not found" }
                )

                if (
                    $task.SendMail.When -notMatch '^Always$|^OnlyWhenFilesAreFound$'
                ) {
                    throw "The value '$($task.SendMail.When)' in 'Tasks.SendMail.When' is not supported. Only the value 'Always' or 'OnlyWhenFilesAreFound' can be used."
                }

                $task.ComputerName | Group-Object |
                Where-Object { $_.Count -ge 2 } |
                ForEach-Object {
                    throw "Property 'ComputerName' contains the duplicate value '$($_.Name)'. Duplicate values are not allowed."
                }

                $task.Filter | Group-Object | Where-Object { $_.Count -ge 2 } |
                ForEach-Object {
                    throw "Property 'Filter' contains the duplicate value '$($_.Name)'. Duplicate values are not allowed."
                }

                $task.FolderPath | Group-Object |
                Where-Object { $_.Count -ge 2 } |
                ForEach-Object {
                    throw "Property 'FolderPath' contains the duplicate value '$($_.Name)'. Duplicate values are not allowed."
                }

                if ($task.PSObject.Properties.Name -notContains 'Recurse') {
                    throw "Property 'Recurse' is mandatory."
                }
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion

        #region Convert .json file
        $tasksToExecute = @()

        foreach ($task in $Tasks) {
            #region Set ComputerName if there is none
            if (
                (-not $task.ComputerName) -or
                ($task.ComputerName -eq 'localhost') -or
                ($task.ComputerName -eq "$env:COMPUTERNAME.$env:USERDNSDOMAIN")
            ) {
                $task.ComputerName = $env:COMPUTERNAME
            }
            #endregion

            #region Create tasks to execute
            foreach ($computerName in $task.ComputerName) {
                foreach ($path in $task.FolderPath) {
                    foreach ($filter in $task.Filter) {
                        $tasksToExecute += [PSCustomObject]@{
                            ComputerName = $computerName
                            Path         = $path
                            Filter       = $filter
                            Recurse      = $task.Recurse
                            SendMail     = $task.SendMail
                            Job          = @{
                                Results = @()
                                Error   = $null
                            }
                        }
                    }
                }
            }
            #endregion
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
        $scriptBlock = {
            try {
                $task = $_

                #region Declare variables for code running in parallel
                if (-not $MaxConcurrentJobs) {
                    $searchScriptPath = $using:searchScriptPath
                    $PSSessionConfiguration = $using:PSSessionConfiguration
                    $EventVerboseParams = $using:EventVerboseParams
                    $EventErrorParams = $using:EventErrorParams
                    $EventOutParams = $using:EventOutParams
                }
                #endregion

                #region Create job parameters
                $invokeParams = @{
                    ComputerName        = $task.ComputerName
                    FilePath            = $searchScriptPath
                    ConfigurationName   = $PSSessionConfiguration
                    ArgumentList        = $task.Path, $task.Filter, $task.Recurse
                    EnableNetworkAccess = $true
                    ErrorAction         = 'Stop'
                }

                $M = "Start job on '{0}' Path '{1}' Filter '{2}' Recurse '{3}'" -f
                $invokeParams.ComputerName,
                $invokeParams.ArgumentList[0],
                $($invokeParams.ArgumentList[1] -join ', '),
                $invokeParams.ArgumentList[2]
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                #endregion

                #region Start job
                $task.Job.Results += Invoke-Command @invokeParams
                #endregion

                #region Verbose
                $M = "Results from '{0}' Path '{1}' Filters '{2}' Recurse '{3}': {4}" -f
                $invokeParams.ComputerName,
                $invokeParams.ArgumentList[0],
                $($invokeParams.ArgumentList[1] -join ', '),
                $invokeParams.ArgumentList[2],
                $task.Job.Results[-1].Count

                if ($errorCount = $task.Job.Results.Where({ $_.Error }).Count) {
                    $M += " , Errors: {0}" -f $errorCount
                    Write-Warning $M
                    Write-EventLog @EventErrorParams -Message $M
                }
                elseif ($task.Job.Results.Count) {
                    Write-Verbose $M
                    Write-EventLog @EventOutParams -Message $M
                }
                else {
                    Write-Verbose $M
                    Write-EventLog @EventVerboseParams -Message $M
                }
                #endregion
            }
            catch {
                $M = "Error on '{0}' Path '{1}' Filters '{2}' Recurse '{3}': $_" -f
                $invokeParams.ComputerName,
                $invokeParams.ArgumentList[0],
                $($invokeParams.ArgumentList[1] -join ', '),
                $invokeParams.ArgumentList[2]
                Write-Warning $M; Write-EventLog @EventErrorParams -Message $M

                $task.Job.Error = $_
                $Error.RemoveAt(0)
            }
        }

        #region Run code serial or parallel
        $foreachParams = if ($MaxConcurrentJobs -eq 1) {
            @{
                Process = $scriptBlock
            }
        }
        else {
            @{
                Parallel      = $scriptBlock
                ThrottleLimit = $MaxConcurrentJobs
            }
        }

        $tasksToExecute | ForEach-Object @foreachParams
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
        #region Count results, errors, ...
        $counter = @{
            totalTasks      = $tasksToExecute.Count
            totalFoundFiles = $tasksToExecute.Job.Results.Count
            executionErrors = $tasksToExecute.where({ $_.Error }).Count
            jobErrors       = $tasksToExecute.Job.where({ $_.Error }).Count
            systemErrors    = ($Error.Exception.Message | Measure-Object).Count
        }
        $counter.totalErrors = $counter.executionErrors +
        $counter.jobErrors + $counter.systemErrors
        #endregion

        #region Exit script when no mail is required
        if (
            (-not $counter.totalFoundFiles) -and
            (-not $counter.totalErrors) -and
            ($file.SendMail.When -ne 'Always')

        ) {
            Exit
        }
        #endregion

        $mailParams = @{
            To        = $file.SendMail.To
            Bcc       = $ScriptAdmin
            LogFolder = $logParams.LogFolder
            Header    = if ($file.SendMail.Header) { $file.SendMail.Header }
            else { $ScriptName }
            Save      = "$logFile - Mail.html"
        }

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'

        $mailParams.Subject = '{0} {1}' -f $counter.sqlFiles, $(
            if ($counter.sqlFiles -ne 1) { 'queries' } else { 'query' }
        )

        if (
            $totalErrorCount = $counter.executionErrors + $counter.jobErrors +
            $counter.systemErrors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -ne 1) { 's' }
            )
        }
        #endregion

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

            $M = "Send mail`r`n- Header:`t{0}`r`n- To:`t`t{1}`r`n- Subject:`t{2}" -f
            $mailParams.Header, $($mailParams.To -join ','),
            $mailParams.Subject
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            Get-ScriptRuntimeHC -Stop
            Send-MailHC @mailParams
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}