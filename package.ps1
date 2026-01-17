# package.ps1
$ErrorActionPreference = "Stop"

$solutionDir = $PSScriptRoot
$appDir = Join-Path $solutionDir "App"
$qtBinDir = "C:\Qt\6.10.1\msvc2022_64\bin" # Adjust if needed

# 1. Clean/Create App Directory
if (Test-Path $appDir) { Remove-Item $appDir -Recurse -Force }
New-Item -ItemType Directory -Path $appDir | Out-Null
Write-Host "Created App directory." -ForegroundColor Cyan

# 2. Check for Release Binaries
# Note: Path still assumes folder structure.
# But binary name is ExplorerSelector.exe
$exePath = Join-Path $solutionDir "ExplorerSelector\x64\Release\ExplorerSelector.exe"
$dllPath = Join-Path $solutionDir "SelectorExplorerPlugin\x64\Release\SelectorExplorerPlugin.dll"

if (-not (Test-Path $exePath)) {
    Write-Warning "Release build of ExplorerSelector.exe not found. Please build Release configuration first."
    Exit
}
if (-not (Test-Path $dllPath)) {
    Write-Warning "Release build of SelectorExplorerPlugin.dll not found. Please build Release configuration first."
    Exit
}

# 3. Copy Binaries
Copy-Item $exePath $appDir
Copy-Item $dllPath $appDir
Write-Host "Copied Binaries." -ForegroundColor Green

# 4. Run windeployqt
$windeployqt = Join-Path $qtBinDir "windeployqt.exe"
if (Test-Path $windeployqt) {
    Write-Host "Running windeployqt..." -ForegroundColor Cyan
    $exeInApp = Join-Path $appDir "ExplorerSelector.exe"
    Start-Process $windeployqt -ArgumentList "--release --no-translations --compiler-runtime `"$exeInApp`"" -NoNewWindow -Wait
} else {
    Write-Warning "windeployqt.exe not found at $qtBinDir. You may need to copy Qt DLLs manually."
}

# 5. Copy Assets
Copy-Item (Join-Path $solutionDir "Assets") $appDir -Recurse
Write-Host "Copied Assets." -ForegroundColor Green

# 6. Process AppxManifest.xml
$manifestSrc = Join-Path $solutionDir "AppxManifest.xml"
$manifestDst = Join-Path $appDir "AppxManifest.xml"
$manifestContent = Get-Content $manifestSrc -Raw

# Update paths: Remove "ExplorerSelector\x64\Debug\" prefix (if any)
$manifestContent = $manifestContent -replace 'Executable=".*\\ExplorerSelector.exe"', 'Executable="ExplorerSelector.exe"'
$manifestContent = $manifestContent -replace 'Path=".*\\SelectorExplorerPlugin.dll"', 'Path="SelectorExplorerPlugin.dll"'

Set-Content $manifestDst $manifestContent
Write-Host "Updated AppxManifest.xml." -ForegroundColor Green

# 7. Process Install.ps1
$installSrc = Join-Path $solutionDir "Install.ps1"
$installDst = Join-Path $appDir "Install.ps1"
$installContent = Get-Content $installSrc -Raw

# Update DLL path logic to be relative to script
$installContent = $installContent -replace 'Join-Path \$projectRoot "SelectorExplorerPlugin\\x64\\Debug\\SelectorExplorerPlugin.dll"', 'Join-Path $projectRoot "SelectorExplorerPlugin.dll"'

Set-Content $installDst $installContent
Write-Host "Updated Install.ps1." -ForegroundColor Green

# 8. Process Uninstall.ps1
$uninstallSrc = Join-Path $solutionDir "Uninstall.ps1"
$uninstallDst = Join-Path $appDir "Uninstall.ps1"
$uninstallContent = Get-Content $uninstallSrc -Raw

# Update DLL path logic
$uninstallContent = $uninstallContent -replace '"\$PSScriptRoot\\SelectorExplorerPlugin\\x64\\Debug\\SelectorExplorerPlugin.dll"', '"$PSScriptRoot\SelectorExplorerPlugin.dll"'

Set-Content $uninstallDst $uninstallContent
Write-Host "Updated Uninstall.ps1." -ForegroundColor Green

Write-Host "Packaging Complete! Content is in $appDir" -ForegroundColor Yellow