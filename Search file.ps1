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
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Search Filter in file name\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)
        
Begin {
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

        try {
            $result = [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                Path         = $Path
                Matches      = @{}
                Error        = $null
            }

            foreach ($filter in $Filters) {
                $params = @{
                    LiteralPath = $Path 
                    Recurse     = $Recurse
                    Filter      = $filter
                    File        = $true
                    ErrorAction = 'Stop'
                }
                $result.Matches[$filter] = Get-ChildItem @params
            }
        }
        catch {
            $result.Error = $_
            $Error.RemoveAt(0)
        }
        finally {
            $result
        }
    }

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

            $task.Filter | Group-Object | Where-Object { $_.Count -ge 2 } |
            ForEach-Object {
                throw "Input file '$ImportFile': Property 'Filter' contains the duplicate value '$($_.Name)'. Duplicate values are not allowed."
            }
            
            foreach ($path in $task.FolderPath) {
                if (-not (Test-Path -LiteralPath $path -PathType Container)) {
                    throw "Input file '$ImportFile': The path '$path' in 'FolderPath' does not exist."
                }
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
            Add-Member -InputObject $task -NotePropertyMembers @{
                Job    = $null
                Result = $null
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
        foreach ($task in $Tasks) {
            foreach ($path in $task.FolderPath) {
                #region Get matching files
                $invokeParams = @{
                    ScriptBlock  = $getMatchingFilesHC
                    ArgumentList = $path, $task.Recurse, $task.Filter
                }
        
                $M = "Search files on '{0}' in folder '{1}' with recurse '{2}' and filter '{3}'" -f $(
                    if ($task.ComputerName) { $task.ComputerName }
                    else { $env:COMPUTERNAME }
                ),
                $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
                $invokeParams.ArgumentList[2]
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        
                $task.Job = if ($task.ComputerName) {
                    $invokeParams.ComputerName = $task.ComputerName
                    $invokeParams.AsJob = $true
                    Invoke-Command @invokeParams
                }
                else {
                    Start-Job @invokeParams
                }
                #endregion

                #region Wait for max running jobs
                $waitParams = @{
                    Name       = $Tasks | Where-Object { $_.Job }
                    MaxThreads = $maxConcurrentJobs
                }
                Wait-MaxRunningJobsHC @waitParams
                #endregion
            }

            if (
                ($task.SendTo.When -eq 'Always') -or
                (1 -eq 1)
            ) {
                
            }
        }

        $totalCount = 0
        
        $tableRows = foreach ($filter in $FileFilters) {
            $params = @{
                Path   = $Path 
                Filter = $filter
                File   = $true
            }
            if ($filesFound = Get-ChildItem @params) {
                $totalCount += $filesFound.Count
        
                '<tr>
                    <th>{0}</th>
                    <td>{1}</td>
                </tr>' -f $filter, $filesFound.Count   
            }
        }

        if ($tableRows) {
            $mailParams = @{
                To        = $MailTo
                Bcc       = $ScriptAdmin
                Priority  = 'Normal'
                Subject   = '{0} errors found' -f $totalCount
                Message   = "
                            <p>Errors found in folder '$Path':</p>
                            <table>
                                $tableRows
                            </table>
                        "
                LogFolder = $logParams.LogFolder
                Header    = $ScriptName
                Save      = $logFile + ' - Mail.html'
                Quote     = $null
            }
            Get-ScriptRuntimeHC -Stop
            Send-MailHC @mailParams
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}