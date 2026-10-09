param([string]$PackageSource)
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
    New-Item -ItemType Directory -Force 'dist' | Out-Null
    Copy-Item -LiteralPath 'src\SheetPilot.Excel\bin\Release\net462\publish\SheetPilot-AddIn-packed.xll' -Destination 'dist\SheetPilot-x86.xll'
    Copy-Item -LiteralPath 'src\SheetPilot.Excel\bin\Release\net462\publish\SheetPilot-AddIn64-packed.xll' -Destination 'dist\SheetPilot-x64.xll'
    foreach ($file in @('README.txt','Check-Environment.ps1','EXCEL-DNA-LICENSE.txt','THIRD-PARTY-NOTICES.txt')) { Copy-Item -LiteralPath $file -Destination 'dist' }
    Get-ChildItem -LiteralPath 'dist' -Filter '*.xll' | Get-FileHash -Algorithm SHA256 | ForEach-Object { $_.Hash + '  ' + [IO.Path]::GetFileName($_.Path) } | Set-Content 'dist\SHA256SUMS.txt' -Encoding ASCII
    Write-Host 'Portable add-ins are ready in dist/.'
} finally { Pop-Location }
