Param (
    [Parameter(Mandatory)]
    [String]$Path,
    [Parameter(Mandatory)]
    [String]$Filter,
    [Parameter(Mandatory)]
    [Boolean]$Recurse
)

try {
    $result = [PSCustomObject]@{
        Files     = @()
        StartTime = Get-Date
        EndTime   = $null
        Error     = $null
    }

    if (-not (Test-Path -LiteralPath $Path -PathType 'Container')) {
        throw "Folder '$Path' not found."
    }

    $params = @{
        LiteralPath = $Path
        Recurse     = $Recurse
        Filter      = $Filter
        File        = $true
        ErrorAction = 'Stop'
    }
    $result.Files += Get-ChildItem @params
}
catch {
    $result.Error = "Failed retrieving files with filter '$Filter': $_"
    $Error.RemoveAt(0)
}
finally {
    $result.EndTime = Get-Date
    $result
}
