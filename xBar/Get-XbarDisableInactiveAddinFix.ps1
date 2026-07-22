# Get-XbarDisableInactiveAddinFix.ps1
# Detects Office version, Xbar ProgID, and generates a .reg file with the needed changes
# Run as the affected user - makes NO system changes
# Admin rights ARE required to import the generated .reg file because HKLM is updated

$knownProgIDs = @("InforCRMXbar", "Saleslogix.Outlook.Connector", "SaleslogixSidebar")

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Xbar Add-in Registry Fix Generator" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# --- Detect Office Version ---
# Trust the highest version found under Outlook\Addins as the active one
# (this will be verified against the actual installed Outlook.exe below)
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

# --- Verify Office version against the actual installed OUTLOOK.EXE ---
# Office uses a shared registry version number (e.g. 16.0 covers Office 2016,
# 2019, 2021, 2024, and Microsoft 365 - they are differentiated by build number,
# not the major version folder). So we only verify the MAJOR version number
# matches the actual installed Outlook.exe, not an exact build match.
# Bitness is also determined directly from the real install path rather than
# assumed, since both 32-bit and 64-bit Outlook can exist on a 64-bit OS.
Write-Host "`n[Verifying against actual installed Outlook.exe]" -ForegroundColor Yellow

$appPathKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"

if (-not (Test-Path $appPathKey)) {
    Write-Host "`n  ERROR: Could not locate OUTLOOK.EXE registration at:" -ForegroundColor Red
    Write-Host "    $appPathKey" -ForegroundColor Red
    Write-Host "  Cannot verify the registry-detected Office version is correct." -ForegroundColor Red
    Write-Host "  Please escalate this output to your Infor CRM administrator." -ForegroundColor Red
    exit 1
}

$outlookExePath = (Get-ItemProperty -Path $appPathKey -ErrorAction SilentlyContinue).'(default)'

if (-not $outlookExePath -or -not (Test-Path $outlookExePath)) {
    Write-Host "`n  ERROR: OUTLOOK.EXE path registered but not found on disk." -ForegroundColor Red
    Write-Host "    Registered Path : $outlookExePath" -ForegroundColor Red
    Write-Host "  Please escalate this output to your Infor CRM administrator." -ForegroundColor Red
    exit 1
}

$outlookFileVersion = (Get-Item $outlookExePath).VersionInfo.ProductVersion
$outlookMajorVersion = ($outlookFileVersion -split '\.')[0] + ".0"

$detectedBitness = if ($outlookExePath -match '\\Program Files \(x86\)\\') { "32-bit" } else { "64-bit" }

Write-Host "  Outlook.exe Path    : $outlookExePath"
Write-Host "  Outlook File Version: $outlookFileVersion"
Write-Host "  Outlook Bitness     : $detectedBitness"
Write-Host "  Registry Version    : $detectedVersion"

if ($outlookMajorVersion -ne $detectedVersion) {
    Write-Host "`n  ERROR: Registry-detected Office version does not match the actual installed Outlook." -ForegroundColor Red
    Write-Host "    Registry Version     : $detectedVersion" -ForegroundColor Red
    Write-Host "    Outlook.exe Path     : $outlookExePath" -ForegroundColor Red
    Write-Host "    Outlook.exe Version  : $outlookFileVersion" -ForegroundColor Red
    Write-Host "  No .reg file will be generated. Please escalate this output to your Infor CRM administrator." -ForegroundColor Red
    exit 1
}

Write-Host "  --> Version confirmed against installed Outlook.exe" -ForegroundColor Green

# --- Find Xbar ProgID ---
# NOTE on 32-bit/64-bit coverage:
# HKCU is NOT subject to Wow6432Node redirection - 32-bit and 64-bit processes
# share the same HKCU hive, so no separate HKCU Wow6432Node path is needed.
# HKLM:\Software\Wow6432Node covers 32-bit add-ins registered machine-wide on
# a 64-bit OS, and is included below alongside the native HKLM path.
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
        # Check both the native CLSID location and the Wow6432Node location,
        # since a 32-bit add-in's CLSID may be registered under either
        # depending on OS/Office bitness combination.
        $dllPath   = $null
        $clsid     = $null
        $dllStatus = "No CLSID found (deeply orphaned)"

        $clsidSearchKeys = @(
            "Registry::HKEY_CLASSES_ROOT\$progID\CLSID",
            "Registry::HKEY_CLASSES_ROOT\Wow6432Node\$progID\CLSID"
        )

        foreach ($clsidKey in $clsidSearchKeys) {
            if (Test-Path $clsidKey) {
                $clsid = (Get-ItemProperty $clsidKey -ErrorAction SilentlyContinue).'(default)'
                if ($clsid) { break }
            }
        }

        if ($clsid) {
            $inprocSearchKeys = @(
                "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\InprocServer32",
                "Registry::HKEY_CLASSES_ROOT\Wow6432Node\CLSID\$clsid\InprocServer32"
            )

            foreach ($inprocKey in $inprocSearchKeys) {
                if (Test-Path $inprocKey) {
                    $dllPath = (Get-ItemProperty $inprocKey -ErrorAction SilentlyContinue).'(default)'
                    if ($dllPath) { break }
                }
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
        Write-Host "    DLL Path : $($c.DllPath)" -ForegroundColor $color
        Write-Host "    DLL Found: $($c.DllExists)" -ForegroundColor $color
    } else {
        Write-Host "    DLL Path : Not resolved" -ForegroundColor Red
        Write-Host "    DLL Found: False" -ForegroundColor Red
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

# --- Check Current State of Target Keys/Values ---
# Reads the four values the .reg file is about to set, so the tech can see
# up front what will actually change vs. what is already correct. This does
# NOT modify anything - read-only.
Write-Host "`n[Checking current state of target registry values]" -ForegroundColor Yellow

$desiredDoNotDisable = 1
$desiredLoadBehavior = 3

function Get-RegDwordValue {
    param(
        [string]$Path,
        [string]$Name
    )
    if (-not (Test-Path $Path)) {
        return [PSCustomObject]@{ Exists = $false; Value = $null }
    }
    $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item -or $null -eq $item.$Name) {
        return [PSCustomObject]@{ Exists = $false; Value = $null }
    }
    return [PSCustomObject]@{ Exists = $true; Value = $item.$Name }
}

$checks = @(
    [PSCustomObject]@{
        Label    = "HKCU DoNotDisableAddinList\$detectedProgID"
        Path     = "HKCU:\Software\Microsoft\Office\$detectedVersion\Outlook\Resiliency\DoNotDisableAddinList"
        Name     = $detectedProgID
        Desired  = $desiredDoNotDisable
    },
    [PSCustomObject]@{
        Label    = "HKLM DoNotDisableAddinList\$detectedProgID"
        Path     = "HKLM:\Software\Microsoft\Office\$detectedVersion\Outlook\Resiliency\DoNotDisableAddinList"
        Name     = $detectedProgID
        Desired  = $desiredDoNotDisable
    },
    [PSCustomObject]@{
        Label    = "HKCU Outlook\Addins\$detectedProgID  LoadBehavior"
        Path     = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$detectedProgID"
        Name     = "LoadBehavior"
        Desired  = $desiredLoadBehavior
    },
    [PSCustomObject]@{
        Label    = "HKLM Outlook\Addins\$detectedProgID  LoadBehavior"
        Path     = "HKLM:\Software\Microsoft\Office\Outlook\Addins\$detectedProgID"
        Name     = "LoadBehavior"
        Desired  = $desiredLoadBehavior
    }
)

$anyChangeNeeded = $false

foreach ($check in $checks) {
    $result = Get-RegDwordValue -Path $check.Path -Name $check.Name

    if (-not $result.Exists) {
        Write-Host "  [MISSING]      $($check.Label)  (not set; will be created as $($check.Desired))" -ForegroundColor Red
        $anyChangeNeeded = $true
    } elseif ($result.Value -eq $check.Desired) {
        Write-Host "  [ALREADY OK]   $($check.Label)  = $($result.Value)" -ForegroundColor Green
    } else {
        Write-Host "  [WILL CHANGE]  $($check.Label)  currently $($result.Value)  ->  will be set to $($check.Desired)" -ForegroundColor Red
        $anyChangeNeeded = $true
    }
}

if ($anyChangeNeeded) {
    Write-Host "`n  --> At least one value differs from the desired state. A .reg file with the necessary fixes will be generated below." -ForegroundColor Yellow
} else {
    Write-Host "`n  --> All target values already match the desired state. A .reg file will still be generated (harmless to reapply), but importing it should make no changes." -ForegroundColor Green
}

# --- Generate .reg file ---
# Both HKCU and HKLM entries are written. Outlook checks both hives for the
# DoNotDisableAddinList and both must be present for full protection against
# the add-in being disabled. HKLM entries require admin rights to import.
$regContent = @"
Windows Registry Editor Version 5.00

; ============================================================
; Xbar Outlook Add-in Fix
; Office Version : $detectedVersion
; Outlook Bitness: $detectedBitness
; ProgID         : $detectedProgID
; Domain         : $domain
; Generated      : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
; ============================================================
; Apply this file by double-clicking it, or via:
;   regedit /s Fix-XbarAddin.reg
; ============================================================

; --- HKCU: Prevent Office from disabling the add-in ---
; HKCU entry is required - Outlook checks both HKCU and HKLM for this list.
[HKEY_CURRENT_USER\Software\Microsoft\Office\$detectedVersion\Outlook\Resiliency\DoNotDisableAddinList]
"$detectedProgID"=dword:00000001

; --- HKCU: Set LoadBehavior to load at startup ---
; NOTE: The Addins path intentionally does NOT include a version number.
; Microsoft documentation and VSTO registry references confirm that LoadBehavior
; is registered under Office\Outlook\Addins, not Office\<version>\Outlook\Addins.
; See: https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins
[HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\$detectedProgID]
"LoadBehavior"=dword:00000003

; ============================================================
; *** IMPORTANT: The entries below require administrator rights.
; If you do not have admin rights, provide this file to your
; IT administrator to apply on your behalf.
; ============================================================

; --- HKLM: Prevent Office from disabling the add-in ---
; HKLM entry is required - Outlook checks both HKCU and HKLM for this list.
; Both must be present for full protection against the add-in being disabled.
[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\$detectedVersion\Outlook\Resiliency\DoNotDisableAddinList]
"$detectedProgID"=dword:00000001

; --- HKLM: Set LoadBehavior to load at startup ---
; NOTE: Same as HKCU above -- no version number in this path by design.
; See: https://learn.microsoft.com/en-us/previous-versions/troubleshoot/outlook/addins-are-registered-under-wow6432node
[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\Addins\$detectedProgID]
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
Write-Host ""
Write-Host "  *** IMPORTANT ***" -ForegroundColor Yellow
Write-Host "  This .reg file contains both HKCU and HKLM entries." -ForegroundColor Yellow
Write-Host "  HKCU entries will apply for the current user without elevation." -ForegroundColor Yellow
Write-Host "  HKLM entries require administrator rights to apply." -ForegroundColor Yellow
Write-Host "  Please provide this file to your IT administrator if you" -ForegroundColor Yellow
Write-Host "  do not have administrator rights on this machine." -ForegroundColor Yellow
Write-Host ""
Write-Host "  Double-click the file to apply, or run: regedit /s Fix-XbarAddin.reg"
Write-Host ""
Write-Host "========================================`n" -ForegroundColor Cyan