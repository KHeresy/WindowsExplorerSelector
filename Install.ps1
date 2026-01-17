# Install.ps1
# Run this script as Administrator

$ErrorActionPreference = "Stop"
$scriptPath = $PSScriptRoot
$projectRoot = $scriptPath

# 1. Register the DLL (Standard COM Registration)
Write-Host "Registering DLL..." -ForegroundColor Cyan
$dllPath = Join-Path $projectRoot "SelectorExplorerPlugin\x64\Debug\SelectorExplorerPlugin.dll"
if (Test-Path $dllPath) {
    Start-Process "regsvr32.exe" -ArgumentList "/s `"$dllPath`"" -Wait
    Write-Host "DLL Registered." -ForegroundColor Green
} else {
    Write-Error "DLL not found at: $dllPath. Please build the project first."
}

# 2. Setup Certificate for Sparse Package
$certName = "ExplorerSelectorDev"
$certSubject = "CN=$certName"
Write-Host "Checking for certificate..." -ForegroundColor Cyan

$cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -eq $certSubject }
if (-not $cert) {
    Write-Host "Creating Self-Signed Certificate..." -ForegroundColor Yellow
    $cert = New-SelfSignedCertificate -Type Custom -Subject $certSubject -KeyUsage DigitalSignature -FriendlyName "ExplorerSelector Dev Cert" -CertStoreLocation "Cert:\LocalMachine\My" -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3", "2.5.29.19={text}")
}

# Trust the certificate
Write-Host "Trusting Certificate..." -ForegroundColor Cyan
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store "TrustedPeople", "LocalMachine"
$store.Open("ReadWrite")
if (-not ($store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint })) {
    $store.Add($cert)
}
$store.Close()

# 3. Register Sparse Package
Write-Host "Registering Sparse Package..." -ForegroundColor Cyan
$manifestPath = Join-Path $projectRoot "AppxManifest.xml"

if (Test-Path $manifestPath) {
    Add-AppxPackage -Register $manifestPath -ForceApplicationShutdown
    Write-Host "Package Registered Successfully!" -ForegroundColor Green
    Write-Host "The context menu should now appear in the top-level Windows 11 menu."
} else {
    Write-Error "AppxManifest.xml not found at $manifestPath"
}