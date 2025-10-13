@ECHO OFF
SET PSSH=C:\Program Files\PowerShell\7\pwsh.exe
SET SCRIPT_PATH=%~dp0..\ExportAllFactTable.ps1
SET CONF_PATH=%~dp0..\config.json

@ECHO ON
"%PSSH%" -ExecutionPolicy Bypass "%SCRIPT_PATH%" -ConfJsonPath "%CONF_PATH%"

@ECHO OFF
SET CONF_PATH=
SET SCRIPT_PATH=
SET PSSH=

@PAUSE
