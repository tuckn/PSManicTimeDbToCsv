$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

function Export-DemensionTable {
<#
.SYNOPSIS
Exports data from a specified table in a SQLite database to a CSV file, filtering records based on a given year and month.

.DESCRIPTION
This PowerShell function, Export-DemensionTable, utilizes SQLite to query data from a specified table within a database. It exports records that match a specific year and month to a CSV file. The function dynamically creates directories for the table and year if they don't already exist, placing the resulting CSV file within these directories. It supports custom paths for the SQLite executable and the destination directory. If not specified, it defaults to using the SQLite command available in the system path and the script's root directory as the destination.

.PARAMETER SqlitePath
The path to the SQLite executable. If not provided, the script assumes `sqlite3` is available in the system path.

.PARAMETER DbPath
Mandatory. The path to the SQLite database file.

.PARAMETER TableName
Mandatory. The name of the table from which to export data.

.PARAMETER DestDirPath
The path to the directory where the CSV file will be saved. Defaults to the script's root directory if not provided.

.EXAMPLE
```powershell
Export-DemensionTable -DbPath "C:\ManicTime\Data\ManicTimeReports.db" -TableName "Ar_Activity" -ColumnName "StartLocalTime" -YearMonth "2024-01" -DestDirPath "C:\exports"
```
This example exports data from the "Activity" table in the "manictime.db" database, filtering records from January 2024. The resulting CSV file is saved in "C:\exports".

This script is intended for users who need to regularly export filtered data from a SQLite database to CSV format, such as for reporting or data analysis purposes. It automates the process of directory creation and file naming based on the table name and specified year and month, streamlining the workflow for recurring exports.
```
#>
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0)]
        [string] $SqlitePath,

        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [String] $DbPath,

        [Parameter(Position = 2, Mandatory = $true)]
        [string] $TableName,

        [Parameter(Position = 3, Mandatory = $true)]
        [string] $ColumnNamesString,

        [Parameter(Position = 3)]
        [String] $DestDirPath
    )
    Process {
        if ([string]::IsNullOrWhiteSpace($SqlitePath)) {
            $SqlitePath = "sqlite3"
        }
        Write-Host "[info] SqlitePath: $SqlitePath"
        Write-Host "[info] DbPath: $DbPath"

        if ([string]::IsNullOrWhiteSpace($DestDirPath)) {
            $DestDirPath = $PSScriptRoot
        }
        elseif (-not (Test-Path -LiteralPath $DestDirPath)) {
            Write-Error "[error] The folder is not existing: `"$DestDirPath`""
            exit 1
        }
        Write-Host "[info] DestDirPath: $DestDirPath"

        [string] $tableDirPath = Join-Path -Path $DestDirPath -ChildPath "Demensions"
        Write-Host "[info] tableDirPath: $tableDirPath"

        # ディレクトリが存在しない場合は作成
        if (-not (Test-Path -Path $tableDirPath)) {
            New-Item -Path $tableDirPath -ItemType Directory
        }

        # 保存するCSVファイルのパスを生成
        [string] $csvFileName = $TableName + ".csv"
        [string] $csvPath = Join-Path -Path $tableDirPath -ChildPath $csvFileName
        Write-Host "[info] csvPath: $csvPath"

        # SQLクエリ
        [string] $sqlQuery = "SELECT $ColumnNamesString FROM $TableName;"
        Write-Host "[info] sqlQuery: $sqlQuery"

        try {
            & $SqlitePath -header -csv $DbPath $sqlQuery > $csvPath
        }
        catch {
            Write-Error $_
            exit 1
        }

        Write-Host "[info] Data exported to CSV at: $csvPath"

        return
    }
}
Export-ModuleMember -Function Export-DemensionTable