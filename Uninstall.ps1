# Uninstall.ps1
# Run this script as Administrator

$ErrorActionPreference = "Continue"

Write-Host "Removing ExplorerFinder Package..." -ForegroundColor Cyan
Get-AppxPackage *ExplorerFinder* | Remove-AppxPackage

Write-Host "Unregistering DLL..." -ForegroundColor Cyan
$dllPath = "$PSScriptRoot\\SelectorExplorerPlugin.dll"
if (Test-Path $dllPath) {
    Start-Process "regsvr32.exe" -ArgumentList "/u /s `"$dllPath`"" -Wait
}

Write-Host "Uninstallation Complete." -ForegroundColor Green

