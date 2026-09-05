# package.ps1
$ErrorActionPreference = "Stop"

$solutionDir = $PSScriptRoot
$appDir = Join-Path $solutionDir "App"
$qtRootDir = $env:QT_ROOT_DIR
if (-not $qtRootDir) {
    throw "QT_ROOT_DIR is not set."
}

$qtBinDir = Join-Path $qtRootDir "bin"

# 1. Clean/Create App Directory
if (Test-Path $appDir) { Remove-Item $appDir -Recurse -Force }
New-Item -ItemType Directory -Path $appDir | Out-Null
Write-Host "Created App directory." -ForegroundColor Cyan

# 2. Check for Release Binaries
# Note: Path still assumes folder structure.
# But binary name is ExplorerSelector.exe
$exePath = Join-Path $solutionDir "x64\Release\ExplorerSelector.exe"
$dllPath = Join-Path $solutionDir "x64\Release\SelectorExplorerPlugin.dll"

if (-not (Test-Path $exePath)) {
    throw "Release build of ExplorerSelector.exe not found: $exePath"
}
if (-not (Test-Path $dllPath)) {
    throw "Release build of SelectorExplorerPlugin.dll not found: $dllPath"
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
    # Allow windeployqt to copy standard Qt translations to "translations" folder
    Start-Process $windeployqt -ArgumentList "--release --compiler-runtime `"$exeInApp`"" -NoNewWindow -Wait
} else {
    throw "windeployqt.exe not found: $windeployqt"
}

# 4.1 Copy Application Translations
# Copy our own translations to the same "translations" folder
$i18nSrc = Join-Path $solutionDir "ExplorerSelector\i18n"
$destTranslationsDir = Join-Path $appDir "translations"

if (Test-Path $i18nSrc) {
    # Ensure destination exists (windeployqt usually creates it, but to be safe)
    if (-not (Test-Path $destTranslationsDir)) {
        New-Item -ItemType Directory -Path $destTranslationsDir -Force | Out-Null
    }
    
    # Copy app-specific translations
    Copy-Item (Join-Path $i18nSrc "explorerselector_*.qm") $destTranslationsDir
    Write-Host "Copied Application Translations to $destTranslationsDir" -ForegroundColor Green
}

# 4.2 Copy Qt Licenses (Best Effort)
$qtRootDir = (Get-Item $qtBinDir).Parent.FullName
$licenseFiles = @("LICENSE.txt", "LICENSE.GPL3-EXCEPT", "LICENSE.LGPLv3", "LICENSE.LGPL3", "LGPL_EXCEPTION.txt")
foreach ($file in $licenseFiles) {
    $srcPath = Join-Path $qtRootDir $file
    if (Test-Path $srcPath) {
        Copy-Item $srcPath $appDir
        Write-Host "Copied Qt License: $file" -ForegroundColor Green
    }
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

Set-Content $manifestDst $manifestContent -Encoding UTF8
Write-Host "Updated AppxManifest.xml." -ForegroundColor Green

# 7. Process Install.bat
$installSrc = Join-Path $solutionDir "Install.bat"
$installDst = Join-Path $appDir "Install.bat"
$installContent = Get-Content $installSrc -Raw

# Update DLL path logic to be relative to script
$installContent = $installContent -replace 'Join-Path \$projectRoot "SelectorExplorerPlugin\\x64\\Debug\\SelectorExplorerPlugin.dll"', 'Join-Path $projectRoot "SelectorExplorerPlugin.dll"'

Set-Content $installDst $installContent
Write-Host "Updated Install.bat." -ForegroundColor Green

# 8. Process UnInstall.bat
$uninstallSrc = Join-Path $solutionDir "UnInstall.bat"
$uninstallDst = Join-Path $appDir "UnInstall.bat"
$uninstallContent = Get-Content $uninstallSrc -Raw

# Update DLL path logic
$uninstallContent = $uninstallContent -replace '"\$PSScriptRoot\\SelectorExplorerPlugin\\x64\\Debug\\SelectorExplorerPlugin.dll"', '"$PSScriptRoot\SelectorExplorerPlugin.dll"'

Set-Content $uninstallDst $uninstallContent
Write-Host "Updated UnInstall.bat." -ForegroundColor Green

Write-Host "Packaging Complete! Content is in $appDir" -ForegroundColor Yellow