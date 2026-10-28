$ErrorActionPreference = "Stop"
$env:COB_MAIN_DIR = "C:\GnuCOBOL"
$env:COB_CONFIG_DIR = "C:\GnuCOBOL\config"
$env:PATH = "C:\GnuCOBOL\bin;" + $env:PATH

cobc -x -free -o calc_pret.exe calc_pret.cbl

if (Test-Path calc_pret.exe) {
    .\calc_pret.exe
}