@ECHO OFF
SET PSSH=C:\Program Files\PowerShell\7\pwsh.exe
SET SCRIPT_PATH=%~dp0..\ExportAllDimensionTable.ps1
SET CONF_PATH=%~dp0..\config.json

@ECHO ON
"%PSSH%" -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -ConfJsonPath "%CONF_PATH%"

@ECHO OFF
SET CONF_PATH=
SET SCRIPT_PATH=
SET PS6_PATH=

@PAUSE
