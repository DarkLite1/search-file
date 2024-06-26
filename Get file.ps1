Param (
    [Parameter(Mandatory)]
    [String]$Path,
    [Parameter(Mandatory)]
    [Boolean]$Recurse,
    [Parameter(Mandatory)]
    [String[]]$Filters
)

if (-not (Test-Path -LiteralPath $Path -PathType 'Container')) {
    throw "Folder '$Path' not found."
}

foreach ($filter in $Filters) {
    try {
        $result = [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            Path         = $Path
            Recurse      = $Recurse
            Filter       = $filter
            Files        = @()
            StartTime    = Get-Date
            EndTime      = $null
            Error        = $null
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
        $result.EndTime = Get-Date
        $result
    }
}