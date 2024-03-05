foreach ($f in Get-ChildItem -Path "$($PSScriptRoot)\Modules\*.psm1") {
    Import-Module -Name $f.FullName -Force
}
