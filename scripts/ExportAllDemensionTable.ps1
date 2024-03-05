using module "..\modules\ExportDemensionTable.psm1"

Param(
    [Parameter(Position = 0)]
    [string] $SqlitePath,

    [Parameter(Position = 1)]
    [String] $DbPath,

    [Parameter(Position = 2)]
    [String] $DestDirPath,

    [Parameter(Position = 3)]
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
    TableName = "Ar_CommonGroup"
    ColumnNamesString = "CommonId, ReportGroupType, KeyHash, GroupType, Key, Name, Color, IsBillable, UpperKey"
    DestDirPath = $DestDirPath
}
Export-DemensionTable @params

$params = @{
    SqlitePath = $SqlitePath
    DbPath = $DbPath
    TableName = "Ar_Group"
    ColumnNamesString = "ReportId, GroupId, ReportGroupType, KeyHash, Key, Name, Color, SkipColor, FolderId, GroupType, IsBillable, CommonId, SourceId, CurrentChangeSequence, CurrentChangeRandomValue, Other"
    DestDirPath = $DestDirPath
}
Export-DemensionTable @params

$params = @{
    SqlitePath = $SqlitePath
    DbPath = $DbPath
    TableName = "Ar_Folder"
    ColumnNamesString = "*"
    DestDirPath = $DestDirPath
}
Export-DemensionTable @params