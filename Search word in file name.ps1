#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Search for files with a matching word in the file name.

    .DESCRIPTION
        This script will report the quantity of files that contain a specific
        key word in their file name.

    .PARAMETER Tasks
        Collection of a items to check.

    .PARAMETER Tasks.FolderPath
        One or more paths to a folder where to look for files.

    .PARAMETER Tasks.MailTo
        List of e-mail addresses where to send the e-mail too.

    .PARAMETER Tasks.Word
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
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Search word in file name\$ScriptName",
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
            if (-not $task.MailTo) {
                throw "Input file '$ImportFile': Property 'MailTo' is mandatory."
            }
            if (-not $task.Word) {
                throw "Input file '$ImportFile': Property 'Word' is mandatory."
            }
            if (-not $task.FolderPath) {
                throw "Input file '$ImportFile': Property 'FolderPath' is mandatory."
            }
            foreach ($path in $task.FolderPath) {
                if (-not (Test-Path -LiteralPath $path -PathType Container)) {
                    throw "Input file '$ImportFile': The path '$path' in 'FolderPath' does not exist."
                }
            }
            if (-not $task.SendMail) {
                throw "Input file '$ImportFile': Property 'SendMail' is mandatory."
            }
            if ($task.SendMail -notMatch '^Always$|^OnlyWhenWordIsFound$') {
                throw "Input file '$ImportFile': The value '$($task.SendMail)' in 'SendMail' is not supported. Only the value 'Always' or 'OnlyWhenWordIsFound' can be used."
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
        $totalCount = 0
        
        $tableRows = foreach ($word in $KeyWords) {
            $params = @{
                Path   = $Path 
                Filter = "*$word*" 
                File   = $true
            }
            if ($filesFound = Get-ChildItem @params) {
                $totalCount += $filesFound.Count
        
                '<tr>
                    <th>{0}</th>
                    <td>{1}</td>
                </tr>' -f $word, $filesFound.Count   
            }
        }

        if ($tableRows) {
            $mailParams = @{
                To        = $MailTo
                Bcc       = $ScriptAdmin
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