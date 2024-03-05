$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

function Export-DemensionTable {
<#
.SYNOPSIS
Exports specified columns from a table in a SQLite database to a CSV file.

.DESCRIPTION
The `Export-DemensionTable` function utilizes the SQLite command-line tool to export data from a specified table within a SQLite database. It allows users to specify which columns to export through a string of column names. The function dynamically creates a directory for the export if it doesn't exist and places the resulting CSV file within this directory. It's designed for flexibility in specifying the SQLite executable path, the database path, table name, columns to export, and the destination directory for the CSV output.

.PARAMETER SqlitePath
The path to the SQLite executable. If not provided, it defaults to `sqlite3`, assuming it's available in the system path.

.PARAMETER DbPath
Mandatory. The path to the SQLite database file from which data will be exported.

.PARAMETER TableName
Mandatory. The name of the table to export data from.

.PARAMETER ColumnNamesString
Mandatory. A string listing the columns to include in the export. Columns should be separated by commas.

.PARAMETER DestDirPath
The directory where the CSV file will be saved. If not provided, it defaults to the script's root directory. If the specified directory does not exist, the function will terminate with an error.

.EXAMPLE
```powershell
Export-DemensionTable -SqlitePath "C:\sqlite\sqlite3.exe" -DbPath "C:\ManicTime\Data\ManicTimeReports.db" -TableName "Ar_CommonGroup" -ColumnNamesString "CommonId, ReportGroupType, KeyHash, GroupType, Key, Name, Color" -DestDirPath "C:\exports"
```
This example exports the `id`, `username`, and `email` columns from the `users` table in the `mydatabase.db` SQLite database. The CSV file is saved in the `C:\exports\Demensions` directory with the name `users.csv`. If the `Demensions` directory does not exist in `C:\exports`, it will be created.

This script streamlines the process of exporting specific table columns to CSV format, making it valuable for data extraction and reporting tasks.
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