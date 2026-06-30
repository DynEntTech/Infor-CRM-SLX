# Get-XbarDisableInactiveAddinFix.ps1
# Detects Office version, Xbar ProgID, and generates a .reg file with the needed changes
# Run as the affected user - makes NO system changes
# Admin rights are NOT required to import the generated .reg file

$knownProgIDs = @("InforCRMXbar", "Saleslogix.Outlook.Connector", "SaleslogixSidebar")

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Xbar Add-in Registry Fix Generator" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# --- Detect Office Version ---
# Trust the highest version found under Outlook\Addins as the active one
$detectedVersion = $null
$detectedProgID  = $null

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
    exit 1
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
    @{ Path = "HKCU:\Software\Microsoft\Office\$detectedVersion\Outlook\Addins"; Hive = "HKCU" },
    @{ Path = "HKLM:\Software\Microsoft\Office\$detectedVersion\Outlook\Addins"; Hive = "HKLM" },
    @{ Path = "HKLM:\Software\Wow6432Node\Microsoft\Office\$detectedVersion\Outlook\Addins"; Hive = "HKLM (Wow6432Node)" }
)

# Collect all candidates across all hives before making any decision
$candidates = @()

foreach ($entry in $addinSearchPaths) {
    if (-not (Test-Path $entry.Path)) { continue }
    Get-ChildItem $entry.Path -ErrorAction SilentlyContinue | ForEach-Object {
        $progID = $_.PSChildName
        $lb     = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).LoadBehavior
        $isXbar = $knownProgIDs | Where-Object { $progID -like "*$_*" }

        if (-not $isXbar) {
            Write-Host "  $progID  |  $($entry.Hive)  |  LoadBehavior: $lb" -ForegroundColor White
            return
        }

        # Resolve DLL path via HKCR: ProgID -> CLSID -> InprocServer32
        $dllPath   = $null
        $clsid     = $null
        $dllStatus = "No CLSID found (deeply orphaned)"

        $clsidKey = "Registry::HKEY_CLASSES_ROOT\$progID\CLSID"
        if (Test-Path $clsidKey) {
            $clsid = (Get-ItemProperty $clsidKey -ErrorAction SilentlyContinue).'(default)'
        }

        if ($clsid) {
            $inprocKey = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\InprocServer32"
            if (Test-Path $inprocKey) {
                $dllPath = (Get-ItemProperty $inprocKey -ErrorAction SilentlyContinue).'(default)'
            }

            if ($dllPath) {
                if (Test-Path $dllPath) {
                    $dllStatus = "DLL Found"
                } else {
                    $dllStatus = "DLL NOT found on disk (orphaned)"
                }
            } else {
                $dllStatus = "CLSID found but no InprocServer32 path (orphaned)"
            }
        }

        $candidates += [PSCustomObject]@{
            ProgID       = $progID
            Hive         = $entry.Hive
            LoadBehavior = $lb
            CLSID        = $clsid
            DllPath      = $dllPath
            DllStatus    = $dllStatus
            DllExists    = ($dllPath -and (Test-Path $dllPath))
        }
    }
}

# Report all Xbar candidates found
Write-Host "`n[Xbar ProgID Candidates]" -ForegroundColor Yellow
if ($candidates.Count -eq 0) {
    Write-Host "`n  ERROR: No Xbar add-in entries were found in the registry." -ForegroundColor Red
    Write-Host "  The following ProgIDs were searched for:" -ForegroundColor Red
    foreach ($id in $knownProgIDs) {
        Write-Host "    - $id" -ForegroundColor Red
    }
    Write-Host "  Please escalate this output to your Infor CRM administrator." -ForegroundColor Red
    exit 1
}

foreach ($c in $candidates) {
    $color = if ($c.DllExists) { "Green" } else { "Red" }
    Write-Host "  $($c.ProgID)  |  $($c.Hive)  |  LoadBehavior: $($c.LoadBehavior)  |  $($c.DllStatus)" -ForegroundColor $color
    if ($c.DllPath) {
        Write-Host "    DLL: $($c.DllPath)" -ForegroundColor $color
    }
}

# Select active ProgID - must have exactly one with a DLL confirmed on disk
$activeCandidates = $candidates | Where-Object { $_.DllExists }

if ($activeCandidates.Count -eq 0) {
    Write-Host "`n  ERROR: No Xbar ProgID could be confirmed with a DLL present on disk." -ForegroundColor Red
    Write-Host "  The following ProgIDs were searched for:" -ForegroundColor Red
    foreach ($id in $knownProgIDs) {
        Write-Host "    - $id" -ForegroundColor Red
    }
    Write-Host "  Please escalate this output to your Infor CRM administrator." -ForegroundColor Red
    exit 1
}

if ($activeCandidates.Count -gt 1) {
    Write-Host "`n  ERROR: More than one Xbar ProgID was found with a DLL present on disk." -ForegroundColor Red
    Write-Host "  The active installation is ambiguous. No .reg file will be generated." -ForegroundColor Red
    Write-Host "  Please escalate this output to your Infor CRM administrator." -ForegroundColor Red
    exit 1
}

$detectedProgID = $activeCandidates[0].ProgID
Write-Host "`n  --> ProgID confirmed: $detectedProgID  |  $($activeCandidates[0].Hive)  |  $($activeCandidates[0].DllPath)" -ForegroundColor Green

# --- Get Domain ---
# Get-CimInstance is preferred in modern PowerShell but requires WS-Man which may
# not be available or configured on older Windows 10 machines. Get-WmiObject is
# used here intentionally for maximum compatibility.
$domain = (Get-WmiObject Win32_ComputerSystem).Domain
Write-Host "`n[Domain]" -ForegroundColor Yellow
Write-Host "  $domain"

# --- Generate .reg file ---
# Only HKCU entries are written. HKCU takes precedence over HKLM for LoadBehavior
# and DoNotDisableAddinList is additive, so HKCU entries are sufficient to fix
# the add-in for the current user regardless of whether Xbar was installed
# machine-wide under HKLM. This avoids requiring admin rights to import the file.
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
; HKCU takes precedence over HKLM for this key. Setting it here is
; sufficient for the current user regardless of how Xbar was installed.
[HKEY_CURRENT_USER\Software\Microsoft\Office\$detectedVersion\Outlook\Resiliency\DoNotDisableAddinList]
"$detectedProgID"=dword:00000001

; --- HKCU: Set LoadBehavior to load at startup ---
; NOTE: The Addins path intentionally does NOT include a version number.
; Microsoft documentation and VSTO registry references confirm that LoadBehavior
; is registered under Office\Outlook\Addins, not Office\<version>\Outlook\Addins.
; See: https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins
[HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\$detectedProgID]
"LoadBehavior"=dword:00000003
"@

# --- Find a writable output location ---
$outputCandidates = @(
    "$env:TEMP",
    "$env:USERPROFILE\Downloads",
    "$env:LOCALAPPDATA\Temp",
    "C:\Temp",
    "$env:USERPROFILE"
)

$outputDir = $null
foreach ($candidate in $outputCandidates) {
    if (Test-Path $candidate) {
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
Write-Host "  Double-click the file to apply, or run: regedit /s Fix-XbarAddin.reg"
Write-Host ""
Write-Host "========================================`n" -ForegroundColor Cyan