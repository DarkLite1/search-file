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
        $jobId = 0

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

            #region Add properties
            $task | Add-Member -NotePropertyMembers @{
                JobId = @()
            }
            #endregion

            #region Create tasks to execute
            foreach ($computerName in $task.ComputerName) {
                foreach ($path in $task.FolderPath) {
                    foreach ($filter in $task.Filter) {
                        $jobId++

                        $task.JobId += $jobId

                        $tasksToExecute += [PSCustomObject]@{
                            ComputerName = $computerName
                            Path         = $path
                            Filter       = $filter
                            Recurse      = $task.Recurse
                            Job          = @{
                                Id      = $jobId
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
        for ($i = 0; $i -lt $Tasks.Count; $i++) {
            $task = $Tasks[$i]

            #region Verbose
            $M = "Task ComputerName '{0}' Path '{1}' Filter '{2}' Recurse '{3}' Mail.To '{4}' Mail.When '{5}'" -f
            $($task.ComputerName -join ', '),
            $($task.FolderPath -join ', '),
            $($task.Filter -join ', '),
            $task.Recurse,
            $($task.SendMail.To -join ', '),
            $task.SendMail.When
            Write-Verbose $M
            #endregion

            $mailParams = @{
                To             = $task.SendMail.To
                Bcc            = $ScriptAdmin
                Priority       = 'Normal'
                LogFolder      = $logParams.LogFolder
                Header         = if ($task.SendMail.Header) {
                    $task.SendMail.Header
                }
                else { $ScriptName }
                EventLogSource = $ScriptName
                Save           = "$logFile - $i - Mail.html"
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

            $executedTasks = $tasksToExecute.Where(
                { $task.JobId -contains $_.Job.Id }
            )

            #region Mail message
            $errorMessage = $null
            $tableRows = @()

            foreach (
                $filters in
                $executedTasks | Group-Object -Property 'ComputerName', 'Path'
            ) {
                $filterRows = foreach ($filter in $filters.Group) {
                    if (
                        ($task.SendMail.When -eq 'Always') -or
                        ($filter.Job.Results)
                    ) {
                        "<tr>
                            <td>{0}</td>
                            <td>{1}</td>
                        </tr>" -f $filter.Filter, $filter.Job.Results.Count
                    }
                }

                if ($filterRows) {
                    $computerName = $filters.Group[0].ComputerName
                    $path = $filters.Group[0].Path

                    $tableRows += "<tr>
                        <th>{0}</th>
                        <th>{1}</th>
                    </tr>
                    <tr>
                        <td>Filter</td>
                        <td>Files found</td>
                    </tr>" -f
                    $computerName,
                    $(
                        if ($path -match '^\\\\') {
                            '<a href="{0}">{0}</a>' -f $path
                        }
                        else {
                            $uncPath = $path -Replace '^.{2}', (
                                '\\{0}\{1}$' -f
                                $computerName, $path[0]
                            )
                            '<a href="{0}">{0}</a>' -f $uncPath
                        }
                    )

                    $tableRows += $filterRows
                }

                $mailParams.Message = "
            $errorMessage
            <p>Found a total of <b>{0} files</b>:</p>
            <table>
                $tableRows
            </table>
            {1}" -f $counter.FilesFound, $(
                    if ($mailParams.Attachments) {
                        '<p><i>* Check the attachment for details</i></p>'
                    }
                )
            }
            #endregion

            #region Create Excel objects
            foreach ($executedTask in $executedTasks) {
                if ($executedTask.Job.Results.Files) {
                    $excelSheet.Files += $executedTask.Job.Results.Files |
                    Select-Object -Property @{
                        Name       = 'ComputerName'
                        Expression = { $executedTask.ComputerName }
                    },
                    @{
                        Name       = 'Folder'
                        Expression = { $_.Directory.FullName }
                    },
                    @{
                        Name       = 'FileName'
                        Expression = { $_.Name }
                    },
                    @{
                        Name       = 'Filter'
                        Expression = { $executedTask.Filter }
                    },
                    @{
                        Name       = 'CreationTime'
                        Expression = { $_.CreationTime }
                    },
                    @{
                        Name       = 'LastWriteTime'
                        Expression = { $_.LastWriteTime }
                    },
                    @{
                        Name       = 'Size'
                        Expression = { [MATH]::Round($_.Length / 1GB, 2) }
                    },
                    @{
                        Name       = 'Size_'
                        Expression = { $_.Length }
                    }
                }

                if ($executedTask.Job.Error) {
                    $excelSheet.Errors += $executedTask.Job.Error | Select-Object -Property @{
                        Name       = 'ComputerName'
                        Expression = { $executedTask.ComputerName }
                    },
                    @{
                        Name       = 'Path'
                        Expression = { $executedTask.Path }
                    },
                    @{
                        Name       = 'Filter'
                        Expression = { $executedTask.Filter }
                    },
                    @{
                        Name       = 'Recurse'
                        Expression = { $executedTask.Recurse }
                    },
                    @{
                        Name       = 'Error'
                        Expression = { $_ }
                    }
                }

                if ($executedTask.Job.Results.Error) {
                    $excelSheet.Errors += $executedTask.Job.Error | Select-Object -Property @{
                        Name       = 'ComputerName'
                        Expression = { $executedTask.ComputerName }
                    },
                    @{
                        Name       = 'Path'
                        Expression = { $executedTask.Path }
                    },
                    @{
                        Name       = 'Filter'
                        Expression = { $executedTask.Filter }
                    },
                    @{
                        Name       = 'Recurse'
                        Expression = { $executedTask.Recurse }
                    },
                    @{
                        Name       = 'Error'
                        Expression = { $_ }
                    }
                }
            }
            #endregion

            #region Counter
            $counter = @{
                FilesFound      = $excelSheet.Files.Count
                ExecutionErrors = $excelSheet.Errors.Count
            }
            #endregion

            #region Create Excel sheets
            if ($excelSheet.Files) {
                $excelParams.WorksheetName = $excelParams.TableName = 'Files'

                $M = "Export {0} rows to sheet '{1}' in Excel file '{2}'" -f
                $counter.FilesFound,
                $excelParams.WorksheetName,
                $excelParams.Path
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

            if ($excelSheet.Errors) {
                $excelParams.WorksheetName = $excelParams.TableName = 'Errors'

                $M = "Export {0} rows to sheet '{1}' in Excel file '{2}'" -f
                $counter.ExecutionErrors,
                $excelParams.WorksheetName,
                $excelParams.Path
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                $excelSheet.Errors | Export-Excel @excelParams

                $mailParams.Attachments = $excelParams.Path
            }
            #endregion

            #region Mail subject and priority
            $mailParams.Subject = '{0} file{1} found' -f
            $counter.FilesFound,
            $(if ($counter.FilesFound -ne 1) { 's' })

            if ($counter.ExecutionErrors) {
                $mailParams.Priority = 'High'

                $mailParams.Subject += ', {0} error{1}' -f
                $counter.ExecutionErrors,
                $(if ($counter.ExecutionErrors -ne 1) { 's' })

                $errorMessage = "<p>Detected <b>{0} error{1}</b> during execution.</p>" -f
                $counter.ExecutionErrors,
                $(if ($counter.ExecutionErrors -ne 1) { 's' })
            }
            #endregion

            #region Check to send mail to user
            $sendMailToUser = $false

            if (
                (
                    $task.SendMail.When -eq 'Always'
                ) -or
                (
                    ($task.SendMail.When -eq 'OnlyWhenFilesAreFound') -and
                    ($counter.FilesFound)
                )
            ) {
                $sendMailToUser = $true
            }
            #endregion

            #region Send mail
            if ($sendMailToUser) {
                Write-Verbose 'Send e-mail to the user'

                if ($counter.ExecutionErrors) {
                    $mailParams.Bcc = $ScriptAdmin
                }

                Send-MailHC @mailParams
            }
            else {
                Write-Verbose 'Send no e-mail to the user'

                if ($counter.ExecutionErrors) {
                    Write-Verbose 'Send e-mail to admin only with errors'

                    $mailParams.To = $ScriptAdmin
                    Send-MailHC @mailParams
                }
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
        Write-EventLog @EventEndParams
    }
}