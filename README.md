# PSManicTimeDbToCsv

Export ManicTime Reports (SQLite) tables to CSV files from PowerShell.

- Targets both fact tables (by month) and dimension tables (full export).
- Bundles `sqlite3.exe` for convenience, or you can point to your own.
- On Windows with Japanese data, use PowerShell 6+ to avoid mojibake (garbled characters).

## Prerequisites

- Windows
- PowerShell 6.0+ (`pwsh`) recommended, especially for Japanese text
- ManicTime Reports DB path (e.g. `C:\Program Files\ManicTime\Data\ManicTimeReports.db`)

## Quick Start (with config.json)

1) Copy the sample and edit paths:

```powershell
Copy-Item .\scripts\config_sample.json .\scripts\config.json
# Edit SqlitePath, DbPath, DestDirPath in .\scripts\config.json
```

2) Run exports:

```powershell
# Export all fact tables for the previous month (default)
pwsh .\scripts\ExportAllFactTable.ps1

# Export all dimension tables
pwsh .\scripts\ExportAllDimensionTable.ps1
```

## Usage Examples (no config file)

You can pass parameters directly instead of using `config.json`.

```powershell
# All fact tables for a specific month
pwsh .\scripts\ExportAllFactTable.ps1 \
  -SqlitePath .\bin\sqlite3.exe \
  -DbPath "C:\\Program Files\\ManicTime\\Data\\ManicTimeReports.db" \
  -DestDirPath "$env:USERPROFILE\\logs\\ManicTime" \
  -YearMonth 2024-07

# All dimension tables
pwsh .\scripts\ExportAllDimensionTable.ps1 \
  -SqlitePath .\bin\sqlite3.exe \
  -DbPath "C:\\Program Files\\ManicTime\\Data\\ManicTimeReports.db" \
  -DestDirPath "$env:USERPROFILE\\logs\\ManicTime"

# One fact table (Ar_Activity for 2024-01)
pwsh .\scripts\ExportFactTable.ps1 \
  -SqlitePath .\bin\sqlite3.exe \
  -DbPath "C:\\Program Files\\ManicTime\\Data\\ManicTimeReports.db" \
  -TableName Ar_Activity \
  -ColumnName StartLocalTime \
  -YearMonth 2024-01 \
  -DestDirPath "$env:USERPROFILE\\logs\\ManicTime"
```

Notes:
- If `-YearMonth` is omitted for fact tables, the previous month is used (format `YYYY-MM`).
- You may omit `-SqlitePath` if `sqlite3` is available in your PATH.

## Configuration File

Default location the scripts look for: `scripts\config.json` (same folder as the `.ps1`).

Schema (see `scripts\config_sample.json`):

```json
{
  "SqlitePath": "C:\\Program Files\\SQLite\\sqlite3.exe",
  "DbPath":  "C:\\Program Files\\ManicTime\\Data\\ManicTimeReports.db",
  "DestDirPath":  "C:\\Users\\YOUR_NAME\\logs\\ManicTime"
}
```

Override location with `-ConfJsonPath`:

```powershell
pwsh .\scripts\ExportAllFactTable.ps1 -ConfJsonPath "D:\\cfg\\manictime.json"
pwsh .\scripts\ExportAllDimensionTable.ps1 -ConfJsonPath "D:\\cfg\\manictime.json"
pwsh .\scripts\ExportFactTable.ps1 -TableName Ar_Activity -ColumnName StartLocalTime -YearMonth 2024-07 -ConfJsonPath "D:\\cfg\\manictime.json"
```

Precedence: Command-line parameters override config values. For any parameter not supplied on the command line, the script uses the value from `config.json`.

## What gets exported

- All Fact Tables (`scripts/ExportAllFactTable.ps1`):
  - `Ar_Activity` (filter by `StartLocalTime`)
  - `Ar_ActivityByHour` (`Hour`)
  - `Ar_ApplicationByDay` (`Hour`)
  - `Ar_ApplicationByYear` (`Hour`)
  - `Ar_DocumentByDay` (`Hour`)
  - `Ar_DocumentByYear` (`Hour`)
  - `Ar_WebSiteByDay` (`Hour`)
  - `Ar_WebSiteByYear` (`Hour`)

- All Dimension Tables (`scripts/ExportAllDimensionTable.ps1`):
  - `Ar_CommonGroup`
  - `Ar_Group`
  - `Ar_Folder`

## Output Structure

- Fact tables: `DestDir\<TableName>\<YYYY>\<MM>.csv`
  - Example: `C:\exports\Ar_Activity\2024\01.csv`
- Dimension tables: `DestDir\Dimensions\<TableName>.csv`
  - Example: `C:\exports\Dimensions\Ar_CommonGroup.csv`

## Avoiding Garbled Japanese Characters (mojibake)

- Use PowerShell 6 or later (`pwsh`) on Windows.
- The scripts redirect the UTF-8 output from `sqlite3` directly to files. Opening in Excel or other tools may require specifying UTF-8.
- If a viewer still shows mojibake, import the CSV specifying UTF‑8 explicitly or open with an editor that defaults to UTF‑8.

## Direct Function Use

Instead of the wrapper scripts, you can import the module and call functions:

```powershell
Import-Module .\ManicTimeDbToCsv.psm1

# Fact example
Export-ManicTimeFactDbToCsv -DbPath "C:\\...\\ManicTimeReports.db" -TableName Ar_Activity -ColumnName StartLocalTime -YearMonth 2024-01 -DestDirPath "C:\\exports"

# Dimension example
Export-DimensionTable -DbPath "C:\\...\\ManicTimeReports.db" -TableName Ar_CommonGroup -ColumnNamesString "CommonId,Name,Color" -DestDirPath "C:\\exports"
```

## License

MIT — see `LICENSE`.

