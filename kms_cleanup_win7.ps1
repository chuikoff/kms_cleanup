# KMS Cleanup Tool — Windows 7 edition
# Requires PowerShell 2.0+, run as Administrator

#Requires -Version 2.0

param(
    [switch]$DryRun,
    [switch]$AutoApprove,
    [string]$LogPath = "$PSScriptRoot\kms_cleanup_win7.log"
)

# =========================
# VERSION
# =========================
$ScriptVersion = "1.0.1"

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

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    return $identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-IsNullOrWhiteSpace {
    param([string]$Value)
    if ($null -eq $Value) { return $true }
    return [string]::IsNullOrEmpty($Value) -or $Value.Trim().Length -eq 0
}

function New-CompatObject {
    param([hashtable]$Properties)
    if ($PSVersionTable.PSVersion.Major -ge 3) {
        return [PSCustomObject]$Properties
    }
    return New-Object PSObject -Property $Properties
}

function Test-CommandAvailable {
    param([string]$Name)
    return $null -ne (Get-Command $Name -ErrorAction SilentlyContinue)
}

function Get-CompatScheduledTasks {
    param([string]$Pattern)

    $found = @()
    $taskRoot = Join-Path $env:SystemRoot 'System32\Tasks'
    if (-not (Test-Path $taskRoot)) { return $found }

    try {
        Log "Scanning task folder (may take a moment)..."
        $items = Get-ChildItem -Path $taskRoot -Recurse -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            if ($item.PSIsContainer) { continue }

            $relative = $item.FullName.Substring($taskRoot.Length)
            $relative = $relative.TrimStart('\')
            $fullName = '\' + $relative
            if ($fullName -match $Pattern -or $item.Name -match $Pattern) {
                $found += New-CompatObject @{
                    DisplayName = $fullName
                    TaskName    = $fullName
                    TaskPath    = $null
                    FullName    = $fullName
                    Method      = 'SchTasks'
                }
            }
        }
    } catch {
        Log "ERROR scanning scheduled tasks: $($_.Exception.Message)"
    }

    return $found
}

function Remove-CompatScheduledTask {
    param($Task)

    $result = & schtasks.exe /delete /tn $Task.FullName /f 2>&1
    if (-not $?) {
        throw "schtasks delete failed: $result"
    }
}

function Get-CompatServices {
    param(
        [string]$Pattern,
        [string[]]$Allowlist
    )

    $services = @()
    if (Test-CommandAvailable 'Get-CimInstance') {
        $services = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue)
    } else {
        $services = @(Get-WmiObject Win32_Service -ErrorAction SilentlyContinue)
    }

    $matched = @()
    foreach ($svc in $services) {
        if ($Allowlist -contains $svc.Name) { continue }
        if ($svc.Name -match $Pattern -or ($svc.PathName -and $svc.PathName -match $Pattern)) {
            $matched += $svc
        }
    }
    return $matched
}

function Get-DefenderService {
    foreach ($name in @('WinDefend', 'MsMpSvc')) {
        $svc = Get-Service $name -ErrorAction SilentlyContinue
        if ($svc) { return $svc }
    }
    return $null
}

function Get-DefenderExclusionRoots {
    return @(
        'HKLM:\SOFTWARE\Microsoft\Windows Defender\Exclusions',
        'HKLM:\SOFTWARE\Microsoft\Microsoft Antimalware\Exclusions'
    )
}

function Get-CompatDefenderExclusions {
    param([string]$Pattern)

    $items = @()

    if (Test-CommandAvailable 'Get-MpPreference') {
        try {
            $def = Get-MpPreference -ErrorAction Stop
            if ($def.ExclusionPath) {
                foreach ($value in $def.ExclusionPath) {
                    if (-not (Test-IsNullOrWhiteSpace $value) -and $value -match $Pattern) {
                        $items += New-CompatObject @{
                            Type         = 'Path'
                            Value        = $value
                            Source       = 'MpPreference'
                            RegistryPath = $null
                        }
                    }
                }
            }
            if ($def.ExclusionProcess) {
                foreach ($value in $def.ExclusionProcess) {
                    if (-not (Test-IsNullOrWhiteSpace $value) -and $value -match $Pattern) {
                        $items += New-CompatObject @{
                            Type         = 'Process'
                            Value        = $value
                            Source       = 'MpPreference'
                            RegistryPath = $null
                        }
                    }
                }
            }
            if ($def.ExclusionExtension) {
                foreach ($value in $def.ExclusionExtension) {
                    if (-not (Test-IsNullOrWhiteSpace $value) -and $value -match $Pattern) {
                        $items += New-CompatObject @{
                            Type         = 'Extension'
                            Value        = $value
                            Source       = 'MpPreference'
                            RegistryPath = $null
                        }
                    }
                }
            }
            return $items
        } catch {
            Log "Get-MpPreference unavailable, using registry fallback: $($_.Exception.Message)"
        }
    }

    $map = @{
        'Paths'     = 'Path'
        'Processes' = 'Process'
        'Extensions' = 'Extension'
    }

    foreach ($root in (Get-DefenderExclusionRoots)) {
        foreach ($subKey in $map.Keys) {
            $keyPath = Join-Path $root $subKey
            if (-not (Test-Path $keyPath)) { continue }

            $key = Get-Item $keyPath
            foreach ($propName in $key.Property) {
                if (-not (Test-IsNullOrWhiteSpace $propName) -and $propName -match $Pattern) {
                    $items += New-CompatObject @{
                        Type         = $map[$subKey]
                        Value        = $propName
                        Source       = 'Registry'
                        RegistryPath = $keyPath
                    }
                }
            }
        }
    }

    return $items
}

function Remove-CompatDefenderExclusion {
    param($Item)

    if ($Item.Source -eq 'MpPreference') {
        switch ($Item.Type) {
            'Path'      { Remove-MpPreference -ExclusionPath $Item.Value -ErrorAction Stop }
            'Process'   { Remove-MpPreference -ExclusionProcess $Item.Value -ErrorAction Stop }
            'Extension' { Remove-MpPreference -ExclusionExtension $Item.Value -ErrorAction Stop }
        }
        return
    }

    Remove-ItemProperty -Path $Item.RegistryPath -Name $Item.Value -ErrorAction Stop
}

function Enable-CompatDefender {
    param($Service, [bool]$HasMpPreference)

    Set-Service $Service.Name -StartupType Automatic -ErrorAction Stop
    Start-Service $Service.Name -ErrorAction Stop

    if ($HasMpPreference) {
        Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction Stop
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
if (-not (Test-IsAdministrator)) {
    Write-Error "Administrator rights required. Run PowerShell as Administrator."
    exit 1
}

$mode = if ($DryRun) { "DRY RUN" } elseif ($AutoApprove) { "AUTO" } else { "INTERACTIVE" }
$os = [System.Environment]::OSVersion.Version
$psVer = $PSVersionTable.PSVersion.ToString()
$taskMethod = 'Tasks folder + schtasks.exe'
$defMethod = if (Test-CommandAvailable 'Get-MpPreference') { 'Get-MpPreference' } else { 'Registry' }

Log "KMS Cleanup Tool for Windows 7 v$ScriptVersion [$mode]"
Log "OS: $($os.Major).$($os.Minor) | PowerShell: $psVer | Tasks: $taskMethod | Defender exclusions: $defMethod"

trap {
    Log "FATAL: $($_.Exception.Message)"
    exit 1
}

# =========================
# AUDIT
# =========================
Section "AUDIT: Scheduled Tasks"

$tasks = @(Get-CompatScheduledTasks -Pattern $Patterns)

if ($tasks.Count -gt 0) {
    $tasks | ForEach-Object {
        Log "Task found: $($_.DisplayName)"
    }
} else {
    Log "No suspicious tasks"
}

Section "AUDIT: Services"

$services = @(Get-CompatServices -Pattern $Patterns -Allowlist $ServiceAllowlist)

if ($services.Count -gt 0) {
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
        if (-not ($prop.Value -is [string])) { continue }
        if ($prop.Value -match $Patterns) {
            $runEntries += New-CompatObject @{
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

$suspiciousExclusions = @(Get-CompatDefenderExclusions -Pattern $Patterns)
$HasMpPreference = Test-CommandAvailable 'Get-MpPreference'

if ($suspiciousExclusions.Count -gt 0) {
    foreach ($item in $suspiciousExclusions) {
        Log "Suspicious $($item.Type.ToLower()) exclusion: $($item.Value) [$($item.Source)]"
    }
} else {
    Log "No suspicious Defender exclusions"
}

Section "AUDIT: Defender Status"

$defSvc = Get-DefenderService

if ($defSvc) {
    Log "Defender: $($defSvc.Name) / $($defSvc.Status) / $($defSvc.StartType)"
} else {
    Log "Defender service missing (WinDefend, MsMpSvc)"
}

# =========================
# REMEDIATION
# =========================
Section "REMEDIATION"

# ---- Tasks ----
if ($tasks.Count -gt 0 -and (Confirm-Action "Remove scheduled tasks?")) {
    foreach ($t in $tasks) {
        Invoke-Remediation "Deleting task: $($t.DisplayName)" {
            Remove-CompatScheduledTask -Task $t
        }
    }
}

# ---- Services ----
if ($services.Count -gt 0 -and (Confirm-Action "Remove services?")) {
    foreach ($s in $services) {
        Invoke-Remediation "Deleting service: $($s.Name)" {
            Stop-Service $s.Name -ErrorAction SilentlyContinue
            $result = & sc.exe delete $s.Name 2>&1
            if (-not $?) {
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
        Invoke-Remediation "Enabling Defender ($($defSvc.Name))" {
            Enable-CompatDefender -Service $defSvc -HasMpPreference $HasMpPreference
        }
    }
}

# ---- Exclusions ----
if ($suspiciousExclusions.Count -gt 0 -and (Confirm-Action "Clear suspicious Defender exclusions?")) {
    foreach ($item in $suspiciousExclusions) {
        Invoke-Remediation "Removing $($item.Type.ToLower()) exclusion: $($item.Value)" {
            Remove-CompatDefenderExclusion -Item $item
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