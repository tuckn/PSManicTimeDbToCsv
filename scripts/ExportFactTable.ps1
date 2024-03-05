using module "..\modules\ExportFactTable.psm1"

Param(
    [Parameter(Position = 0)]
    [string] $SqlitePath,

    [Parameter(Position = 1)]
    [String] $DbPath,

    [Parameter(Position = 2)]
    [string] $TableName,

    [Parameter(Position = 3)]
    [string] $ColumnName,

    [Parameter(Position = 4)]
    [string] $YearMonth,

    [Parameter(Position = 5)]
    [String] $DestDirPath,

    [Parameter(Position = 6)]
    [String] $ConfJsonPath
)

Set-StrictMode -Version 3.0

# Set config.json
if ([string]::IsNullOrWhiteSpace($ConfJsonPath)) {
    $ConfJsonPath = Join-Path $PSScriptRoot "config.json"
}
Write-Host "ConfJsonPath: $ConfJsonPath"

if (Test-Path -LiteralPath $ConfJsonPath) {
    $configVals = Get-Content -Path $ConfJsonPath | ConvertFrom-Json

    if ([string]::IsNullOrWhiteSpace($SqlitePath)) {
        $SqlitePath = $configVals.SqlitePath
    }

    if ([string]::IsNullOrWhiteSpace($DbPath)) {
        $DbPath = $configVals.DbPath
    }

    if ([string]::IsNullOrWhiteSpace($DestDirPath)) {
        $DestDirPath = $configVals.DestDirPath
    }
}

$params = @{
    SqlitePath = $SqlitePath
    DbPath = $DbPath
    TableName = $TableName
    ColumnName = $ColumnName
    YearMonth = $YearMonth
    DestDirPath = $DestDirPath
}

Export-ManicTimeFactDbToCsv @params