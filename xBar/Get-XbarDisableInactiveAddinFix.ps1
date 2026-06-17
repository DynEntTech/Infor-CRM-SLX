# Get-XbarDisableInactiveAddinFix.ps1
# Detects Office version, Xbar ProgID, and generates a .reg file with the needed changes
# Run as the affected user - makes NO system changes
# Admin rights are required to import the generated .reg file because HKCM is updated

$knownProgIDs = @("InforCRMXbar", "Saleslogix.Outlook.Connector", "SaleslogixSidebar")

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Xbar Add-in Registry Fix Generator" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# --- Detect Office Version ---
# Trust the highest version found under Outlook\Addins as the active one
$detectedVersion = $null
$detectedProgID  = $null
$detectedHive    = $null

$searchPaths = @(
    @{ Hive = "HKCU"; Root = "HKCU:\Software\Microsoft\Office" },
    @{ Hive = "HKLM"; Root = "HKLM:\Software\Microsoft\Office" },
    @{ Hive = "HKLM (Wow6432Node)"; Root = "HKLM:\Software\Wow6432Node\Microsoft\Office" }
)

$versionsFound = @{}

foreach ($entry in $searchPaths) {
    if (-not (Test-Path $entry.Root)) { continue }
    Get-ChildItem $entry.Root -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | ForEach-Object {
        $ver = $_.PSChildName
        $addinsPath = "$($entry.Root)\$ver\Outlook\Addins"
        if (Test-Path $addinsPath) {
            if (-not $versionsFound[$ver]) { $versionsFound[$ver] = @() }
            $versionsFound[$ver] += $entry.Hive
        }
    }
}

Write-Host "`n[Office Versions with Outlook Addins key]" -ForegroundColor Yellow
if ($versionsFound.Count -eq 0) {
    Write-Host "  None found." -ForegroundColor Red
} else {
    foreach ($ver in ($versionsFound.Keys | Sort-Object -Descending)) {
        Write-Host "  $ver  ->  found in: $($versionsFound[$ver] -join ', ')"
    }
    $detectedVersion = ($versionsFound.Keys | Sort-Object -Descending | Select-Object -First 1)
    Write-Host "  --> Using version: $detectedVersion" -ForegroundColor Green
}

# --- Find Xbar ProgID ---
Write-Host "`n[Searching for Xbar ProgID under version $detectedVersion]" -ForegroundColor Yellow

$addinSearchPaths = @(
    "HKCU:\Software\Microsoft\Office\$detectedVersion\Outlook\Addins",
    "HKLM:\Software\Microsoft\Office\$detectedVersion\Outlook\Addins",
    "HKLM:\Software\Wow6432Node\Microsoft\Office\$detectedVersion\Outlook\Addins"
)

foreach ($path in $addinSearchPaths) {
    if (-not (Test-Path $path)) { continue }
    Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
        $progID = $_.PSChildName
        $lb     = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).LoadBehavior
        $isXbar = $knownProgIDs | Where-Object { $progID -like "*$_*" }
        $tag    = if ($isXbar) { " <-- Xbar" } else { "" }
        Write-Host "  $progID  |  LoadBehavior: $lb$tag" -ForegroundColor $(if ($isXbar) { "Green" } else { "White" })
        if ($isXbar -and -not $detectedProgID) {
            $detectedProgID = $progID
            $detectedHive   = $path -replace '\\Software.*', ''
        }
    }
}

if (-not $detectedProgID) {
    Write-Host "  No known Xbar ProgID found. Please check the list above and re-run with the correct ProgID." -ForegroundColor Red
    exit 1
}

Write-Host "`n  --> ProgID: $detectedProgID" -ForegroundColor Green

# --- Get Domain ---
$domain = (Get-WmiObject Win32_ComputerSystem).Domain
Write-Host "`n[Domain]" -ForegroundColor Yellow
Write-Host "  $domain"

# --- Generate .reg file ---
# Convert HKCU/HKLM to reg file format (no colons, backslash paths)
$regContent = @"
Windows Registry Editor Version 5.00

; ============================================================
; Xbar Outlook Add-in Fix
; Office Version : $detectedVersion
; ProgID         : $detectedProgID
; Domain         : $domain
; Generated      : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
; ============================================================
; Apply this file by double-clicking it, or via:
;   regedit /s Fix-XbarAddin.reg
; ============================================================

; --- HKCU: Prevent Office from disabling the add-in ---
[HKEY_CURRENT_USER\Software\Microsoft\Office\$detectedVersion\Outlook\Resiliency\DoNotDisableAddinList]
"$detectedProgID"=dword:00000001

; --- HKCU: Set LoadBehavior to load at startup ---
[HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\$detectedProgID]
"LoadBehavior"=dword:00000003

; --- HKLM: Prevent Office from disabling the add-in (for locked-down accounts) ---
[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\$detectedVersion\Outlook\Resiliency\DoNotDisableAddinList]
"$detectedProgID"=dword:00000001

; --- HKLM: Set LoadBehavior to load at startup (for locked-down accounts) ---
[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\Addins\$detectedProgID]
"LoadBehavior"=dword:00000003
"@

# --- Find a writable output location (hidden from casual view) ---
$candidates = @(
    "$env:TEMP",
    "$env:USERPROFILE\Downloads",
    "$env:LOCALAPPDATA\Temp",
    "C:\Temp",
    "$env:USERPROFILE"
)

$outputDir = $null
foreach ($candidate in $candidates) {
    if (Test-Path $candidate) {
        # Verify we can actually write there
        $testFile = Join-Path $candidate "xbar_write_test.tmp"
        try {
            [IO.File]::WriteAllText($testFile, "test")
            Remove-Item $testFile -Force
            $outputDir = $candidate
            break
        } catch {
            # Not writable, try next
        }
    }
}

if (-not $outputDir) {
    Write-Host "  Could not find a writable output location." -ForegroundColor Red
    exit 1
}

$outputPath = Join-Path $outputDir "Fix-XbarAddin.reg"
$regContent | Out-File -FilePath $outputPath -Encoding ASCII

Write-Host "`n[Output]" -ForegroundColor Yellow
Write-Host "  .reg file saved to: $outputPath" -ForegroundColor Green
Write-Host "  Review it, then double-click to apply (requires admin for HKLM entries)."
Write-Host ""
Write-Host "========================================`n" -ForegroundColor Cyan
