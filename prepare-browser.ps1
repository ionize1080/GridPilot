param([string]$Destination = (Join-Path $PSScriptRoot '.temp/Browser'))
$ErrorActionPreference='Stop'
$version='44.7.0'
$sha='EEE30DC8FA1F5EA95490E59F44E46EA68DD24C6E93D22FACF70FE5C2D4C2665C'
$archive=Join-Path $PSScriptRoot '.temp/electron.zip'
New-Item -ItemType Directory -Force (Split-Path $archive) | Out-Null
if (-not (Test-Path $archive)) { Invoke-WebRequest "https://github.com/electron/electron/releases/download/v$version/electron-v$version-win32-x64.zip" -OutFile $archive }
if ((Get-FileHash $archive -Algorithm SHA256).Hash -ne $sha) { throw 'Chromium runtime checksum mismatch.' }
Expand-Archive -LiteralPath $archive -DestinationPath $Destination -Force
Write-Host "Portable Chromium runtime prepared. No installer was run. Electron $version (Windows x64)."
