param(
    [switch]$DryRun,
    [switch]$AutoApprove,
    [string]$LogPath = "$PSScriptRoot\kms_cleanup.log"
)

# =========================
# ENCODING FIX
# =========================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# =========================
# LOGGER
# =========================
function Log {
    param([string]$msg)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

# =========================
# CONFIRM
# =========================
function Confirm {
    param([string]$msg)
    if ($AutoApprove) { return $true }
    $ans = Read-Host "$msg (y/n)"
    return $ans -eq "y"
}

# =========================
# SLMGR WRAPPER
# =========================
function Run-Slmgr {
    param([string]$args)

    $output = cscript.exe $env:SystemRoot\System32\slmgr.vbs $args 2>&1

    if ($output -match "0xC004D302") {
        Log "slmgr $args -> skipped (rearm limit reached)"
    }
    else {
        Log "slmgr $args -> executed"
    }
}

# =========================
# SECTION HEADER
# =========================
function Section {
    param([string]$name)
    Log ""
    Log "=============================="
    Log $name
    Log "=============================="
}

# =========================
# CONFIG
# =========================
$Patterns = "KMS|AutoKMS|AAct|Pico"

# =========================
# AUDIT START
# =========================
Section "AUDIT: Scheduled Tasks"

$tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
    $_.TaskName -match $Patterns
}

if ($tasks) {
    $tasks | ForEach-Object {
        Log "Task found: $($_.TaskName)"
    }
} else {
    Log "No suspicious tasks"
}

Section "AUDIT: Services"

$services = Get-CimInstance Win32_Service | Where-Object {
    $_.Name -match $Patterns -or $_.PathName -match $Patterns
}

if ($services) {
    $services | ForEach-Object {
        Log "Service found: $($_.Name) | $($_.PathName)"
    }
} else {
    Log "No suspicious services"
}

Section "AUDIT: Run Keys"

$runHits = @()

$runPaths = @(
"HKLM:\Software\Microsoft\Windows\CurrentVersion\Run",
"HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
)

foreach ($p in $runPaths) {
    if (Test-Path $p) {
        $props = Get-ItemProperty $p
        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Value -match $Patterns) {
                $runHits += "$($prop.Name) => $($prop.Value)"
            }
        }
    }
}

if ($runHits.Count -gt 0) {
    $runHits | ForEach-Object { Log "Run entry: $_" }
} else {
    Log "No Run entries"
}

Section "AUDIT: Defender Exclusions"

$def = Get-MpPreference

if ($def.ExclusionPath -or $def.ExclusionProcess -or $def.ExclusionExtension) {
    $def.ExclusionPath | ForEach-Object { Log "Path exclusion: $_" }
    $def.ExclusionProcess | ForEach-Object { Log "Process exclusion: $_" }
    $def.ExclusionExtension | ForEach-Object { Log "Extension exclusion: $_" }
} else {
    Log "No Defender exclusions"
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
if ($tasks -and (Confirm "Remove scheduled tasks?")) {
    foreach ($t in $tasks) {
        Log "Deleting task: $($t.TaskName)"
        if (-not $DryRun) {
            Unregister-ScheduledTask -TaskName $t.TaskName -Confirm:$false
        }
    }
}

# ---- Services ----
if ($services -and (Confirm "Remove services?")) {
    foreach ($s in $services) {
        Log "Deleting service: $($s.Name)"
        if (-not $DryRun) {
            Stop-Service $s.Name -Force -ErrorAction SilentlyContinue
            sc.exe delete $s.Name | Out-Null
        }
    }
}

# ---- Defender ----
if ($defSvc) {
    if ($defSvc.Status -eq "Running") {
        Log "Defender OK -> skip enable"
    } else {
        if (Confirm "Enable Defender?") {
            try {
                if (-not $DryRun) {
                    Set-Service WinDefend -StartupType Automatic
                    Start-Service WinDefend
                    Set-MpPreference -DisableRealtimeMonitoring $false
                }
            } catch {
                Log "Defender blocked (Tamper Protection likely)"
            }
        }
    }
}

# ---- Exclusions ----
if (($def.ExclusionPath -or $def.ExclusionProcess -or $def.ExclusionExtension) -and (Confirm "Clear Defender exclusions?")) {
    if (-not $DryRun) {
        $def.ExclusionPath | ForEach-Object { Remove-MpPreference -ExclusionPath $_ }
        $def.ExclusionProcess | ForEach-Object { Remove-MpPreference -ExclusionProcess $_ }
        $def.ExclusionExtension | ForEach-Object { Remove-MpPreference -ExclusionExtension $_ }
    }
}

# ---- Activation ----
if (Confirm "Reset Windows activation?") {
    if (-not $DryRun) {
        Run-Slmgr "/upk"
        Run-Slmgr "/ckms"
        Run-Slmgr "/rearm"
    }
}

Section "DONE"
Log "Reboot recommended"
