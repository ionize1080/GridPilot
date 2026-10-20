$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    New-Item -ItemType Directory -Force .temp | Out-Null
    $installer = Join-Path $PSScriptRoot '.temp/Chatbox-1.20.3-Setup.exe'
    if (-not (Test-Path -LiteralPath $installer)) {
        Invoke-WebRequest 'https://github.com/chatboxai/chatbox/releases/download/v1.20.3/Chatbox-1.20.3-Setup.exe' -OutFile $installer
    }
    if ((Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash -ne '874171C6C7358FF3915F01412EB036154A5DED7A76C4663FB0F592D2C5C23B13') {
        throw 'Official Chatbox 1.20.3 checksum mismatch.'
    }
    & npm ci --prefix integrations/chatbox --ignore-scripts --no-audit --no-fund
    if ($LASTEXITCODE -ne 0) { throw 'Chatbox build tool restore failed.' }
    & node integrations/chatbox/prepare.cjs
    if ($LASTEXITCODE -ne 0) { throw 'Chatbox preparation failed.' }
} finally { Pop-Location }
