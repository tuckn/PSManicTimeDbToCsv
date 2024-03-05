$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

function Export-ManicTimeFactDbToCsv {
<#
.SYNOPSIS
Exports data from a specified table in a SQLite database to a CSV file, filtering records based on a given year and month.

.DESCRIPTION
This PowerShell function, Export-ManicTimeFactDbToCsv, utilizes SQLite to query data from a specified table within a database. It exports records that match a specific year and month to a CSV file. The function dynamically creates directories for the table and year if they don't already exist, placing the resulting CSV file within these directories. It supports custom paths for the SQLite executable and the destination directory. If not specified, it defaults to using the SQLite command available in the system path and the script's root directory as the destination.

.PARAMETER SqlitePath
The path to the SQLite executable. If not provided, the script assumes `sqlite3` is available in the system path.

.PARAMETER DbPath
Mandatory. The path to the SQLite database file.

.PARAMETER TableName
Mandatory. The name of the table from which to export data.

.PARAMETER ColumnName
Mandatory. The name of the column used to filter records by year and month.

.PARAMETER YearMonth
The year and month used to filter records, in "YYYY-MM" format. Defaults to the previous month if not specified.

.PARAMETER DestDirPath
The path to the directory where the CSV file will be saved. Defaults to the script's root directory if not provided.

.EXAMPLE
```powershell
Export-ManicTimeFactDbToCsv -DbPath "C:\ManicTime\Data\ManicTimeReports.db" -TableName "Ar_Activity" -ColumnName "StartLocalTime" -YearMonth "2024-01" -DestDirPath "C:\exports"
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
        [string] $ColumnName,

        [Parameter(Position = 4)]
        [string] $YearMonth,

        [Parameter(Position = 5)]
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

        # $YearMonthが空欄の場合、先月の年月を計算して格納
        if ([string]::IsNullOrWhiteSpace($YearMonth)) {
            $lastMonth = (Get-Date).AddMonths(-1)
            $YearMonth = $lastMonth.ToString("yyyy-MM")
        }
        Write-Host "[info] YearMonth: $YearMonth"

        # フォルダと年4桁のフォルダを作成
        [string] $tableDirPath = Join-Path -Path $DestDirPath -ChildPath $TableName
        [string] $yearDirPath = Join-Path -Path $tableDirPath -ChildPath $YearMonth.Substring(0,4)

        Write-Host "[info] tableDirPath: $tableDirPath"
        Write-Host "[info] yearDirPath: $yearDirPath"

        # ディレクトリが存在しない場合は作成
        if (-not (Test-Path -Path $tableDirPath)) {
            New-Item -Path $tableDirPath -ItemType Directory
        }

        if (-not (Test-Path -Path $yearDirPath)) {
            New-Item -Path $yearDirPath -ItemType Directory
        }

        # 保存するCSVファイルのパスを生成
        [string] $csvFileName = $YearMonth.Substring(5,2) + ".csv"
        [string] $csvPath = Join-Path -Path $yearDirPath -ChildPath $csvFileName
        Write-Host "[info] csvPath: $csvPath"

        # SQLクエリ
        [string] $sqlQuery = "SELECT * FROM $TableName WHERE strftime('%Y-%m', $ColumnName) = '$YearMonth';"
        Write-Host "[info] sqlQuery: $sqlQuery"

        try {
            # @FIXME!
            # 様々な方法を試したが、日本語のWindows環境で文字化けせずにCSVファイルを出力することができなかった…。

            # [A]: SQLクエリの定義し流し込む方法は、改行を認識せず？動作しない
            # $sqlQuery = @"
# .mode csv
# .output '$csvPath'
# SELECT *
# FROM $TableName
# WHERE strftime('%Y-%m', $ColumnName) = '$YearMonth';
# "@
            # & $SqlitePath $DbPath $sqlQuery

            # [B]: 一時ファイルに出力してから、UTF-8 BOM付きで出力ファイルに書き出す方法も試したが、うまくいかなかった…。
            # & $SqlitePath -header -csv $DbPath $sqlQuery > $tempFile
            # $content = Get-Content -Path $tempPath -Raw
            # [System.IO.File]::WriteAllText($csvPath, $content, [System.Text.Encoding]::UTF8)
            # Remove-Item $tempPath

            # [C]: SQLiteからの出力をOut-FileでUTF8を指定しても文字化けした
            # [string] $tempFile = [System.IO.Path]::GetTempFileName()
            # Write-Host "[info] tempFile: $tempFile"
            # & $SqlitePath -header -csv $DbPath $sqlQuery | Out-File -FilePath $tempFile -Encoding UTF8

            # [D]: パイプを使わず$outputに格納する方法も試したが、この時点の$outputで文字化けしている
            # $output = & $SqlitePath -header -csv $DbPath $sqlQuery
            # $utf8EncodingWithBOM = New-Object System.Text.UTF8Encoding $true
            # [System.IO.File]::WriteAllText($csvPath, $output, $utf8EncodingWithBOM)

            # [E]: Cと同じだが、こちらはSystem.IO.StreamWriterをつかった
            # $output = & $SqlitePath -header -csv $DbPath $sqlQuery
            # $writer = [System.IO.StreamWriter]::new($csvPath, $false, [System.Text.Encoding]::UTF8)
            # $writer.Write($output)
            # $writer.Close()

            # [F]: Start-Processで標準出力をリダイレクトする方法も試したが、Error: incomplete inputが発生
            # $process = Start-Process -FilePath $SqlitePath -ArgumentList "-header", "-csv", $DbPath, $sqlQuery -NoNewWindow -RedirectStandardOutput $csvPath -PassThru
            # $process.WaitForExit()

            # [G]: [B]と同じ。だがPowerShell 6以上必要なら文字化けしなかった
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
Export-ModuleMember -Function Export-ManicTimeFactDbToCsv