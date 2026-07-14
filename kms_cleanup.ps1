#Requires -RunAsAdministrator

param(
    [switch]$DryRun,
    [switch]$AutoApprove,
    [string]$LogPath = "$PSScriptRoot\kms_cleanup.log"
)

# =========================
# VERSION
# =========================
$ScriptVersion = "1.2.0"

# =========================
# ENCODING FIX
# =========================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# =========================
# LOGGER
# =========================
function Log {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $Message"
    Write-Host $line
    try {
        Add-Content -Path $LogPath -Value $line -ErrorAction Stop
    } catch {
        Write-Warning "Failed to write log: $($_.Exception.Message)"
    }
}

# =========================
# CONFIRM
# =========================
function Confirm-Action {
    param([string]$Message)
    if ($AutoApprove -or $DryRun) { return $true }
    $ans = Read-Host "$Message (y/n)"
    return $ans -match '^[yY]'
}

# =========================
# SLMGR WRAPPER
# =========================
function Run-Slmgr {
    param([string]$SlmgrArgs)

    try {
        $output = cscript.exe $env:SystemRoot\System32\slmgr.vbs $SlmgrArgs 2>&1 | Out-String

        if ($output -match "0xC004D302") {
            Log "slmgr $SlmgrArgs -> skipped (rearm limit reached)"
        } else {
            Log "slmgr $SlmgrArgs -> executed"
        }
    } catch {
        Log "ERROR running slmgr $SlmgrArgs : $($_.Exception.Message)"
    }
}

# =========================
# SECTION HEADER
# =========================
function Section {
    param([string]$Name)
    Log ""
    Log "=============================="
    Log $Name
    Log "=============================="
}

function Invoke-Remediation {
    param(
        [string]$Action,
        [scriptblock]$ScriptBlock
    )

    $prefix = if ($DryRun) { "[DRY RUN] " } else { "" }
    Log "$prefix$Action"

    if (-not $DryRun) {
        try {
            & $ScriptBlock
        } catch {
            Log "ERROR: $($_.Exception.Message)"
        }
    }
}

# =========================
# CONFIG
# =========================
$Patterns = 'AutoKMS|AAct(?:Net)?|KMSAuto|KMSSS|HEU[_\s]?KMS|Pico|Office\d+KMS|VLC[_\s]?KMS|KMSpico|KMSPico|KMS@Net'
$ServiceAllowlist = @('SPPKMSvc')

# =========================
# START
# =========================
$mode = if ($DryRun) { "DRY RUN" } elseif ($AutoApprove) { "AUTO" } else { "INTERACTIVE" }
Log "KMS Cleanup Tool v$ScriptVersion [$mode]"

# =========================
# AUDIT
# =========================
Section "AUDIT: Scheduled Tasks"

$tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
    $_.TaskName -match $Patterns
}

if ($tasks) {
    $tasks | ForEach-Object {
        Log "Task found: $($_.TaskPath)$($_.TaskName)"
    }
} else {
    Log "No suspicious tasks"
}

Section "AUDIT: Services"

$services = Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -notin $ServiceAllowlist -and (
        $_.Name -match $Patterns -or $_.PathName -match $Patterns
    )
}

if ($services) {
    $services | ForEach-Object {
        Log "Service found: $($_.Name) | $($_.PathName)"
    }
} else {
    Log "No suspicious services"
}

Section "AUDIT: Run Keys"

$runEntries = @()

$runPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
)

foreach ($path in $runPaths) {
    if (-not (Test-Path $path)) { continue }

    $props = Get-ItemProperty $path
    foreach ($prop in $props.PSObject.Properties) {
        if ($prop.Name -like 'PS*') { continue }
        if ($prop.Value -isnot [string]) { continue }
        if ($prop.Value -match $Patterns) {
            $runEntries += [PSCustomObject]@{
                Path  = $path
                Name  = $prop.Name
                Value = $prop.Value
            }
            Log "Run entry: $($prop.Name) => $($prop.Value) [$path]"
        }
    }
}

if ($runEntries.Count -eq 0) {
    Log "No Run entries"
}

Section "AUDIT: Defender Exclusions"

$def = $null
try {
    $def = Get-MpPreference -ErrorAction Stop
} catch {
    Log "Cannot read Defender preferences: $($_.Exception.Message)"
}

$suspiciousPaths = @()
$suspiciousProcesses = @()
$suspiciousExtensions = @()

if ($def) {
    if ($def.ExclusionPath) {
        $suspiciousPaths = @($def.ExclusionPath | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_) -and $_ -match $Patterns
        })
    }
    if ($def.ExclusionProcess) {
        $suspiciousProcesses = @($def.ExclusionProcess | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_) -and $_ -match $Patterns
        })
    }
    if ($def.ExclusionExtension) {
        $suspiciousExtensions = @($def.ExclusionExtension | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_) -and $_ -match $Patterns
        })
    }

    if ($suspiciousPaths) { $suspiciousPaths | ForEach-Object { Log "Suspicious path exclusion: $_" } }
    if ($suspiciousProcesses) { $suspiciousProcesses | ForEach-Object { Log "Suspicious process exclusion: $_" } }
    if ($suspiciousExtensions) { $suspiciousExtensions | ForEach-Object { Log "Suspicious extension exclusion: $_" } }

    if (-not $suspiciousPaths -and -not $suspiciousProcesses -and -not $suspiciousExtensions) {
        Log "No suspicious Defender exclusions"
    }
} else {
    Log "Defender exclusions audit skipped"
}

Section "AUDIT: Defender Status"

$defSvc = Get-Service WinDefend -ErrorAction SilentlyContinue

if ($defSvc) {
    Log "Defender: $($defSvc.Status) / $($defSvc.StartType)"
} else {
    Log "Defender service missing"
}

# =========================
# REMEDIATION
# =========================
Section "REMEDIATION"

# ---- Tasks ----
if ($tasks -and (Confirm-Action "Remove scheduled tasks?")) {
    foreach ($t in $tasks) {
        Invoke-Remediation "Deleting task: $($t.TaskPath)$($t.TaskName)" {
            Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction Stop
        }
    }
}

# ---- Services ----
if ($services -and (Confirm-Action "Remove services?")) {
    foreach ($s in $services) {
        Invoke-Remediation "Deleting service: $($s.Name)" {
            Stop-Service $s.Name -Force -ErrorAction SilentlyContinue
            $result = sc.exe delete $s.Name 2>&1
            if ($LASTEXITCODE -ne 0) {
                throw "sc.exe delete failed: $result"
            }
        }
    }
}

# ---- Run keys ----
if ($runEntries.Count -gt 0 -and (Confirm-Action "Remove Run entries?")) {
    foreach ($entry in $runEntries) {
        Invoke-Remediation "Removing Run entry: $($entry.Name) from $($entry.Path)" {
            Remove-ItemProperty -Path $entry.Path -Name $entry.Name -ErrorAction Stop
        }
    }
}

# ---- Defender ----
if ($defSvc) {
    if ($defSvc.Status -eq "Running" -and $defSvc.StartType -eq "Automatic") {
        Log "Defender OK -> skip"
    } elseif (Confirm-Action "Enable Defender?") {
        Invoke-Remediation "Enabling Defender" {
            Set-Service WinDefend -StartupType Automatic -ErrorAction Stop
            Start-Service WinDefend -ErrorAction Stop
            if ($def) {
                Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction Stop
            }
        }
    }
}

# ---- Exclusions ----
$hasSuspiciousExclusions = $suspiciousPaths -or $suspiciousProcesses -or $suspiciousExtensions

if ($hasSuspiciousExclusions -and $def -and (Confirm-Action "Clear suspicious Defender exclusions?")) {
    foreach ($p in $suspiciousPaths) {
        Invoke-Remediation "Removing path exclusion: $p" {
            Remove-MpPreference -ExclusionPath $p -ErrorAction Stop
        }
    }

    foreach ($p in $suspiciousProcesses) {
        Invoke-Remediation "Removing process exclusion: $p" {
            Remove-MpPreference -ExclusionProcess $p -ErrorAction Stop
        }
    }

    foreach ($e in $suspiciousExtensions) {
        Invoke-Remediation "Removing extension exclusion: $e" {
            Remove-MpPreference -ExclusionExtension $e -ErrorAction Stop
        }
    }
}

# ---- Activation ----
if (Confirm-Action "Reset Windows activation?") {
    if ($DryRun) {
        Log "[DRY RUN] Would run: slmgr /upk, /ckms, /rearm"
    } else {
        Run-Slmgr "/upk"
        Run-Slmgr "/ckms"
        Run-Slmgr "/rearm"
    }
}

Section "DONE"
Log "Reboot recommended"