param([string]$PackageSource, [string]$BrowserRuntime)
$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    foreach ($project in @('src\SheetPilot.Excel\SheetPilot.Excel.csproj','tests\SheetPilot.Tests\SheetPilot.Tests.csproj')) {
        $restoreArgs = @('restore', $project, '--locked-mode')
        if ($PackageSource) { $restoreArgs += @('--source', $PackageSource) }
        & dotnet @restoreArgs
        if ($LASTEXITCODE -ne 0) { throw 'Dependency restore failed.' }
        & dotnet build $project -c Release --no-restore
        if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
    }
    & '.\tests\SheetPilot.Tests\bin\Release\net462\SheetPilot.Tests.exe'
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
    & '.\tests\SheetPilot.Tests\bin\Release\net462\SheetPilot.Tests.exe' --color-picker-host
    if ($LASTEXITCODE -ne 0) { throw 'Cold native-host color picker tests failed.' }
    & '.\tests\SheetPilot.Tests\bin\Release\net462\SheetPilot.Tests.exe' --release-features
    if ($LASTEXITCODE -ne 0) { throw 'Release feature tests failed.' }
    foreach($suite in @('--navigation-host','--upgrade','--v1','--v11','--multi-sort','--ai')) {
        & '.\tests\SheetPilot.Tests\bin\Release\net462\SheetPilot.Tests.exe' $suite
        if ($LASTEXITCODE -ne 0) { throw "Regression tests failed: $suite" }
    }
    New-Item -ItemType Directory -Force 'dist' | Out-Null
    Copy-Item -LiteralPath 'src\SheetPilot.Excel\bin\Release\net462\publish\SheetPilot-AddIn-packed.xll' -Destination 'dist\GridPilot-x86.xll'
    Copy-Item -LiteralPath 'src\SheetPilot.Excel\bin\Release\net462\publish\SheetPilot-AddIn64-packed.xll' -Destination 'dist\GridPilot-x64.xll'
    Get-ChildItem 'src/SheetPilot.Excel/bin/Release/net462/*.dll' | Copy-Item -Destination 'dist' -Force
    Copy-Item -LiteralPath 'App' -Destination 'dist' -Recurse -Force
    if (-not (Test-Path '.temp/Chatbox/bootstrap.cjs')) { throw 'Run prepare-chatbox.ps1 to prepare the pinned official Chatbox 1.20.3 bundle.' }
    Copy-Item -LiteralPath '.temp/Chatbox' -Destination 'dist' -Recurse -Force
    if (-not $BrowserRuntime) { $BrowserRuntime=Join-Path $PSScriptRoot '.temp/Browser' }
    if (-not (Test-Path (Join-Path $BrowserRuntime 'electron.exe'))) { throw 'Run prepare-browser.ps1 or provide -BrowserRuntime. No runtime installation is needed.' }
    $previousElectronMode=$env:ELECTRON_RUN_AS_NODE
    try {
        $env:ELECTRON_RUN_AS_NODE='1'
        foreach($suite in @('tests/reconcile.cjs','tests/reconcile-manager.cjs','tests/ai-tools.cjs')) {
            $testProcess=Start-Process -FilePath (Join-Path $BrowserRuntime 'electron.exe') -ArgumentList $suite -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -Wait -PassThru
            if ($testProcess.ExitCode -ne 0) { throw "Reconciliation tests failed: $suite" }
        }
    } finally { $env:ELECTRON_RUN_AS_NODE=$previousElectronMode }
    New-Item -ItemType Directory -Force 'dist/Browser' | Out-Null
    Get-ChildItem -LiteralPath $BrowserRuntime | Copy-Item -Destination 'dist/Browser' -Recurse -Force
    Copy-Item -LiteralPath 'src\SheetPilot.Core\bin\Release\net462\Resources' -Destination 'dist' -Recurse -Force
    Copy-Item -LiteralPath 'README.md','Check-Environment.ps1' -Destination 'dist' -Force
    Copy-Item -LiteralPath 'docs' -Destination 'dist' -Recurse -Force
    $packageRoot=(Resolve-Path dist).Path
    Get-ChildItem -LiteralPath $packageRoot -File -Recurse | Where-Object Name -ne 'SHA256SUMS.txt' | Sort-Object FullName | Get-FileHash -Algorithm SHA256 | ForEach-Object { $_.Hash + '  ' + $_.Path.Substring($packageRoot.Length+1).Replace('\','/') } | Set-Content 'dist\SHA256SUMS.txt' -Encoding ASCII
    Write-Host 'Portable add-ins are ready in dist/.'
} finally { Pop-Location }
