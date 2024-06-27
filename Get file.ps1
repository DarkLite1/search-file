﻿Param (
    [Parameter(Mandatory)]
    [String]$Path,
    [Parameter(Mandatory)]
    [String[]]$Filters,
    [Parameter(Mandatory)]
    [Boolean]$Recurse
)

if (-not (Test-Path -LiteralPath $Path -PathType 'Container')) {
    throw "Folder '$Path' not found."
}

foreach ($filter in $Filters) {
    try {
        $result = [PSCustomObject]@{
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