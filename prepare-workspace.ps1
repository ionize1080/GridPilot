param([string]$BuildPython='python')
$ErrorActionPreference='Stop'
Push-Location $PSScriptRoot
try {
    New-Item -ItemType Directory -Force '.temp','dist/Python' | Out-Null
    $archive=Join-Path $PSScriptRoot '.temp/python-3.12.10.zip'
    if (-not (Test-Path -LiteralPath $archive)) {
        Invoke-WebRequest 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip' -OutFile $archive
    }
    if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ne '4ACBED6DD1C744B0376E3B1CF57CE906F9DC9E95E68824584C8099A63025A3C3') { throw 'Python runtime checksum mismatch.' }
    Expand-Archive -LiteralPath $archive -DestinationPath 'dist/Python' -Force
    & $BuildPython -m pip install --only-binary=:all: --platform win_amd64 --python-version 3.12 --implementation cp --abi cp312 --target 'dist/Python/Lib/site-packages' -r requirements-workspace-lock.txt
    if ($LASTEXITCODE -ne 0) { throw 'Workspace dependency preparation failed.' }
    @('python312.zip','.','Lib\site-packages','import site') | Set-Content 'dist/Python/python312._pth' -Encoding ASCII
    Copy-Item -LiteralPath 'bridge' -Destination 'dist' -Recurse -Force
} finally { Pop-Location }
