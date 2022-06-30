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
        One or more paths to a folder where to look for files.

    .PARAMETER Tasks.MailTo
        List of e-mail addresses where to send the e-mail too.

    .PARAMETER Tasks.Filter
        One or more strings to search for in the file name.

    .PARAMETER Tasks.MailOnlyWhenFound
        When set to true an e-mail will only be sent when matching file 
        names are found.
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
        Function Get-MatchingFilesHC {
            Param (
                [Parameter(Mandatory)]
                [String]$Path,
                [Parameter(Mandatory)]
                [String[]]$Filters
            )

            try {
                $result = [PSCustomObject]@{
                    Path    = $Path
                    Matches = @{}
                    Error   = $null
                }

                foreach ($filter in $Filters) {
                    $params = @{
                        LiteralPath = $Path 
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

        foreach ($task in $Tasks) {
            
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