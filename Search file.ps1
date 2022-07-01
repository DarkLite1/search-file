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

        $M = "'{0}' job duration '{1:hh}:{1:mm}:{1:ss}:{1:fff}'" -f 
        $Job.Location, $jobDuration
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

        #region Get job results
        $M = "'{0}' job get results" -f $Job.Location
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
            $M = "'{0}' job error '{1}'" -f $Job.Location, $e.ToString()
            Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
                  
            $result.Errors += $M
            $error.Remove($e)
        }
        if ($result.Result.Error) {
            $M = "'{0}' error '{1}'" -f $Job.Location, $result.Result.Error
            Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
   
            $result.Errors += $M
        }
        #endregion

        $result.Result = $result.Result | 
        Select-Object -Property * -ExcludeProperty 'Error'

        if (-not $result.Errors) {
            $M = "'{0}' job successful" -f $Job.Location
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

        try {
            $result = [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                Path         = $Path
                Matches      = @{}
                Error        = $null
            }

            if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
                throw "Path '$path' not found."
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
    $getJobResult = {
        #region Verbose
        $M = "Get job results on ComputerName '{0}' Path '{1}' Filter '{2}' Recurse '{3}'" -f $(
            if ($task.ComputerName) { $task.ComputerName }
            else { $env:COMPUTERNAME }
        ), 
        $completedJob.Path, $task.Filter, $task.Recurse
        Write-Verbose $M
        Write-EventLog @EventVerboseParams -Message $M
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
            $task | Add-Member -NotePropertyMembers @{
                Jobs = foreach ($path in $task.FolderPath) {
                    [PSCustomObject]@{
                        Path = $path
                        Job  = @{
                            Object   = $null
                            Duration = $null
                            Result   = $null
                            Errors   = @()
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
        
                $M = "Start job on ComputerName '{0}' Path '{1}' Filter '{3}' Recurse '{2}'" -f $(
                    if ($task.ComputerName) { $task.ComputerName }
                    else { $env:COMPUTERNAME }
                ),
                $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
                $invokeParams.ArgumentList[2]
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        
                $j.Job.Object = if ($task.ComputerName) {
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
                    Name       = $Tasks.Jobs.Job.Object | Where-Object { $_ }
                    MaxThreads = $maxConcurrentJobs
                }
                Wait-MaxRunningJobsHC @waitParams
                #endregion

                #region Get job results
                foreach ($task in $Tasks) {
                    foreach (
                        $completedJob in 
                        $task.Jobs.Job | Where-Object {
                            ($_.Object.State -match 'Completed|Failed')
                        }
                    ) {
                        & $getJobResult
                    }
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

            foreach ($task in $Tasks) {
                $completedJob = $task.Jobs | Where-Object {
                    ($_.Job.Object.Id -eq $finishedJob.Id)
                }
                & $getJobResult
                break
            }
        }
        #endregion

        foreach ($task in $Tasks) {

            #region Export job results to Excel file
            if ($jobResults = $task.Jobs.Job.Result | Where-Object { $_ }) {
                $M = "Export $($jobResults.Count) rows to Excel"
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
                $excelParams = @{
                    Path               = $logFile + ' - Log.xlsx'
                    WorksheetName      = 'Overview'
                    TableName          = 'Overview'
                    NoNumberConversion = '*'
                    AutoSize           = $true
                    FreezeTopRow       = $true
                }
                $jobResults | 
                Select-Object -Property * -ExcludeProperty 'PSComputerName',
                'RunSpaceId', 'PSShowComputerName', 'Output' | 
                Export-Excel @excelParams

                $mailParams.Attachments = $excelParams.Path
            }
            #endregion
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