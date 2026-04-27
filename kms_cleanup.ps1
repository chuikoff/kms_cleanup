param(
    [switch]$DryRun,
    [switch]$AutoApprove,
    [string]$LogPath = "$PSScriptRoot\kms_cleanup.log"
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

function Confirm-Step($msg) {
    if ($AutoApprove) { return $true }
    $a = Read-Host "$msg (y/n)"
    return $a -eq "y"
}

function Section($name) {
    Log ""
    Log "=== $name ==="
}

function Run-Slmgr($args) {
    $output = cscript.exe $env:SystemRoot\System32\slmgr.vbs $args 2>&1

    if ($output -match "0xC004D302") {
        Log "slmgr $args → skipped (rearm limit reached)"
    } else {
        Log "slmgr $args → done"
    }
}

# =========================
# AUDIT
# =========================

$Report = @{}

Section "Audit: Scheduled Tasks"
$tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
    $_.TaskName -match "KMS|AutoKMS|AAct"
}
$Report.Tasks = $tasks
if ($tasks) { $tasks | % { Log "Task: $($_.TaskName)" } } else { Log "OK" }

Section "Audit: Services"
$services = Get-Service -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -match "KMS|AAct"
}
$Report.Services = $services
if ($services) { $services | % { Log "Service: $($_.Name)" } } else { Log "OK" }

Section "Audit: Run keys"
$runFindings = @()
$paths = @(
 "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run",
 "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
)

foreach ($p in $paths) {
    if (Test-Path $p) {
        $props = Get-ItemProperty $p
        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Value -match "KMS|AAct") {
                $runFindings += "$($prop.Name) => $($prop.Value)"
            }
        }
    }
}
$Report.Run = $runFindings
if ($runFindings) { $runFindings | % { Log "Run: $_" } } else { Log "OK" }

Section "Audit: Defender exclusions"
$def = Get-MpPreference
if ($def.ExclusionPath -or $def.ExclusionProcess -or $def.ExclusionExtension -or $def.ExclusionIpAddress) {
    $def.ExclusionPath | % { Log "Path: $_" }
    $def.ExclusionProcess | % { Log "Process: $_" }
    $def.ExclusionExtension | % { Log "Ext: $_" }
    $def.ExclusionIpAddress | % { Log "IP: $_" }
} else {
    Log "OK"
}

Section "Audit: Defender status"
$defSvc = Get-Service WinDefend -ErrorAction SilentlyContinue
$Report.DefenderService = $defSvc

if ($defSvc) {
    Log "Status=$($defSvc.Status) Startup=$($defSvc.StartType)"
} else {
    Log "Not found"
}

Section "Audit finished"

# =========================
# REMEDIATION
# =========================

Section "Remediation"

if ($Report.Tasks -and (Confirm-Step "Delete scheduled tasks?")) {
    foreach ($t in $Report.Tasks) {
        Log "Deleting task: $($t.TaskName)"
        if (-not $DryRun) {
            Unregister-ScheduledTask -TaskName $t.TaskName -Confirm:$false
        }
    }
}

if ($Report.Services -and (Confirm-Step "Delete services?")) {
    foreach ($s in $Report.Services) {
        Log "Deleting service: $($s.Name)"
        if (-not $DryRun) {
            Stop-Service $s.Name -Force -ErrorAction SilentlyContinue
            sc.exe delete $s.Name | Out-Null
        }
    }
}

if ($Report.Run -and (Confirm-Step "Clean Run keys?")) {
    foreach ($entry in $Report.Run) {
        Log "Found: $entry"
    }
}

if (($def.ExclusionPath -or $def.ExclusionProcess -or $def.ExclusionExtension -or $def.ExclusionIpAddress) -and (Confirm-Step "Clear Defender exclusions?")) {
    if (-not $DryRun) {
        $def.ExclusionPath | % { Remove-MpPreference -ExclusionPath $_ }
        $def.ExclusionProcess | % { Remove-MpPreference -ExclusionProcess $_ }
        $def.ExclusionExtension | % { Remove-MpPreference -ExclusionExtension $_ }
        $def.ExclusionIpAddress | % { Remove-MpPreference -ExclusionIpAddress $_ }
    }
}

# Defender — только если реально нужно
if ($Report.DefenderService) {
    if ($Report.DefenderService.Status -eq "Running" -and $Report.DefenderService.StartType -eq "Automatic") {
        Log "Defender OK — skipping"
    } else {
        if (Confirm-Step "Enable Defender?")) {
            try {
                if (-not $DryRun) {
                    Set-Service WinDefend -StartupType Automatic
                    Start-Service WinDefend
                }
            } catch {
                Log "Access denied (Tamper Protection?)"
            }
        }
    }
}

if (Confirm-Step "Reset Windows activation?") {
    if (-not $DryRun) {
        Run-Slmgr "/upk"
        Run-Slmgr "/ckms"
        Run-Slmgr "/rearm"
    }
}

Log "=== DONE. Reboot recommended ==="
