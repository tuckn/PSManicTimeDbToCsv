$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

function Export-ManicTimeDbToCsv {
    # 引数から年月とCSVファイルの保存パスを受け取る
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0)]
        [string] $SqlitePath,

        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [String] $DbPath,

        [Parameter(Position = 2)]
        [string] $YearMonth,

        [Parameter(Position = 3, Mandatory = $true)]
        [string] $TableName,

        [Parameter(Position = 4)]
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

        # SQLクエリの定義
        $sqlQuery = @"
.mode csv
.output '$csvPath'
SELECT *
FROM $TableName
WHERE strftime('%Y-%m', StartLocalTime) = '$YearMonth';
"@

        Write-Host "[info] sqlQuery: $sqlQuery"

        try {
            # sqlite3コマンドの実行
            # sqlite3.exeのパスが環境変数に含まれている必要があります
            # そうでない場合は、sqlite3のフルパスを指定してください（例: C:\sqlite\sqlite3.exe）
            & $SqlitePath $DbPath $sqlQuery
        }
        catch {
            Write-Error $_
            exit 1
        }

        Write-Host "[info] Data exported to CSV at: $csvPath"

        return
    }
}
Export-ModuleMember -Function Export-ManicTimeDbToCsv