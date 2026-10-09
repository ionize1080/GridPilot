param([string]$ExcelPath)
$ErrorActionPreference = 'Stop'
Write-Host 'SheetPilot environment check (read-only)'
$release = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
if ($release -ge 394802) { Write-Host '.NET Framework 4.6.2 or later: OK' }
else { Write-Warning '.NET Framework 4.6.2 or later is required. No software will be installed by this script.' }
if (-not $ExcelPath) {
    foreach ($key in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe')) {
        $item = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
        if ($item) { $candidate = $item.'(default)'; if ($candidate -and (Test-Path -LiteralPath $candidate)) { $ExcelPath = $candidate; break } }
    }
}
if (-not $ExcelPath) {
    foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not $base) { continue }
        foreach ($sub in @('Microsoft Office\Office14\EXCEL.EXE','Microsoft Office\root\Office16\EXCEL.EXE','Microsoft Office\Office16\EXCEL.EXE','Microsoft Office\Office15\EXCEL.EXE')) {
            $candidate = Join-Path $base $sub
            if (Test-Path -LiteralPath $candidate) { $ExcelPath = $candidate; break }
        }
        if ($ExcelPath) { break }
    }
}
if (-not $ExcelPath) { Write-Warning 'Excel path not found. Check Excel Help/About for bitness, or run with -ExcelPath.'; exit 1 }
$stream = [IO.File]::OpenRead($ExcelPath)
$reader = [IO.BinaryReader]::new($stream)
try {
    if ($reader.ReadUInt16() -ne 0x5A4D) { throw 'Not a Windows executable.' }
    $stream.Position = 0x3c; $offset = $reader.ReadInt32(); $stream.Position = $offset
    if ($reader.ReadUInt32() -ne 0x4550) { throw 'Invalid PE header.' }
    $machine = $reader.ReadUInt16()
} finally { $reader.Dispose(); $stream.Dispose() }
Write-Host ('Excel: ' + $ExcelPath)
Write-Host ('Version: ' + (Get-Item -LiteralPath $ExcelPath).VersionInfo.ProductVersion)
switch ($machine) {
    0x014c { Write-Host 'Use SheetPilot-x86.xll (32-bit Excel).' }
    0x8664 { Write-Host 'Use SheetPilot-x64.xll (64-bit Excel).' }
    default { Write-Warning ('Unsupported Excel machine type: ' + $machine) }
}
Write-Host 'No installation, registry changes, or Excel launch was performed.'
