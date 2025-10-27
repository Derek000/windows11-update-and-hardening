#requires -RunAsAdministrator
<#
.SYNOPSIS
  One-command Windows 11 update and sensible hardening.
.DESCRIPTION
  - Updates Windows (quality + Microsoft Update), Winget sources & apps, Windows Defender signatures
  - Optional: DISM/SFC health repair, WSL update, optional ASR rules
  - Conservative defaults, detailed logging (file + Windows Event Log), supports -WhatIf
  - Interactive progress indicators with spinner & step banners (disable via -NoProgress)
.NOTES
  Aligned with OWASP, NIST CSF, CIS Controls. Safe defaults. Use switches for stronger hardening.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [switch]$All,                  # run everything (updates + health + minimal hardening)
  [switch]$UpdateOnly,           # only updates, no hardening
  [switch]$MinimalHardening,     # apply conservative Defender preferences (default if -All)
  [switch]$ApplyASR,             # apply recommended ASR rules (may impact some workflows)
  [switch]$SkipWinget,
  [switch]$SkipWindowsUpdate,
  [switch]$SkipDefender,
  [switch]$SkipHealthRepair,     # skip DISM/SFC
  [switch]$SkipWSL,
  [switch]$NoReboot,             # do not reboot automatically
  [switch]$RebootIfNeeded,       # reboot automatically if pending
  [switch]$VerboseLog,           # more verbose console/file logging
  [switch]$NoProgress,           # suppress progress UI/spinner
  [string]$LogDir                # optional custom log directory
)

# ----------------------------- Globals & Logging -----------------------------
$script:StartTime = Get-Date
$script:LogRoot = if ($PSBoundParameters.ContainsKey('LogDir')) { $LogDir } else { Join-Path $env:ProgramData "Madrock\Update-Strengthen\logs" }
$baseName = "run_{0:yyyyMMdd_HHmmss}" -f $script:StartTime
$script:LogPath = Join-Path $script:LogRoot ($baseName + "_ops.log")               # our ops log
$script:TranscriptPath = Join-Path $script:LogRoot ($baseName + "_transcript.log") # PowerShell transcript
$script:EventSource = "Madrock.UpdateStrengthen"
$script:SpinnerState = @{} # id -> {Timer, I, Activity, Status}
$ErrorActionPreference = 'Stop'

function New-LogInfra {
  New-Item -ItemType Directory -Path $script:LogRoot -Force | Out-Null
  try {
    if (-not [System.Diagnostics.EventLog]::SourceExists($script:EventSource)) {
      New-EventLog -LogName Application -Source $script:EventSource
    }
  } catch { }
  try { Start-Transcript -Path $script:TranscriptPath -Append | Out-Null } catch {}
}

function Write-Log {
  param(
    [Parameter(Mandatory)][string]$Message,
    [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')] [string]$Level = 'INFO'
  )
  $prefix = "[{0}] {1}" -f $Level, (Get-Date).ToString("u")
  $line = "$prefix $Message"

  if ($VerboseLog -or $Level -in @('WARN','ERROR','DEBUG')) { Write-Host $line }

  try {
    $fs = [System.IO.File]::Open($script:LogPath,
      [System.IO.FileMode]::OpenOrCreate,
      [System.IO.FileAccess]::Write,
      [System.IO.FileShare]::ReadWrite)
    $sw = New-Object System.IO.StreamWriter($fs)
    $sw.BaseStream.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
    $sw.WriteLine($line)
    $sw.Flush()
    $sw.Dispose()
    $fs.Dispose()
  } catch {
    Write-Warning "Log write failed: $($_.Exception.Message)"
  }

  $eventId = switch($Level){ 'ERROR' {3} 'WARN' {2} 'SUCCESS' {1} default {0} }
  try { Write-EventLog -LogName Application -Source $script:EventSource -EventId (1000 + $eventId) -EntryType Information -Message $line } catch {}
}

function Stop-LogInfra {
  try { Stop-Transcript | Out-Null } catch {}
}

function Assert-Admin {
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) { throw "Please run as Administrator." }
}

# ----------------------------- Progress UI -----------------------------
function Start-Spinner {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Activity,
    [Parameter(Mandatory)][string]$Status,
    [int]$Id = 1
  )
  if ($NoProgress) { return }
  # Timer-based progress updates in current runspace
  $timer = New-Object Timers.Timer
  $timer.Interval = 250
  $script:SpinnerState[$Id] = @{ Timer = $timer; I = 0; Activity = $Activity; Status = $Status }
  Register-ObjectEvent -InputObject $timer -EventName Elapsed -MessageData $Id -Action {
    $id = $event.MessageData
    $s = $script:SpinnerState[$id]
    if ($null -eq $s) { return }
    $i = $s.I + 3
    $script:SpinnerState[$id].I = $i
    $p = ($i % 100)
    Write-Progress -Id $id -Activity $s.Activity -Status $s.Status -PercentComplete $p
  } | Out-Null
  $timer.Start()
}

function Stop-Spinner {
  [CmdletBinding()]
  param([int]$Id = 1)
  if ($NoProgress) { return }
  if ($script:SpinnerState.ContainsKey($Id)) {
    $timer = $script:SpinnerState[$Id].Timer
    try { $timer.Stop(); $timer.Dispose() } catch {}
    Write-Progress -Id $Id -Activity $script:SpinnerState[$Id].Activity -Status "Completed" -Completed
    $script:SpinnerState.Remove($Id) | Out-Null
  }
}

function Banner {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Text)
  $line = ('=' * 78)
  Write-Host "`n$line`n$Text`n$line`n"
}

function Invoke-Step {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Name,
    [Parameter(Mandatory)][ScriptBlock]$Action,
    [int]$Id = 1,
    [string]$Hint,
    [switch]$Long
  )
  $hintText = if ($Long) {
    if ([string]::IsNullOrEmpty($Hint)) {
      "This step downloads/installs updates and can appear to pause while packages are processed."
    } else { $Hint }
  } else { $Hint }

  Write-Log "$Name - starting."
  if (-not $NoProgress) {
    Write-Host ("[{0}] {1}" -f ((Get-Date).ToString("HH:mm:ss")), $Name)
    if ($hintText) { Write-Host "  > $hintText" }
  }
  Start-Spinner -Id $Id -Activity "Windows Update & Strengthen" -Status "$Name — $hintText"
  $sw = [Diagnostics.Stopwatch]::StartNew()
  try {
    & $Action
    $sw.Stop()
    Write-Log "$Name - completed in $([int]$sw.Elapsed.TotalSeconds)s." 'SUCCESS'
  } catch {
    $sw.Stop()
    Write-Log "$Name - error after $([int]$sw.Elapsed.TotalSeconds)s: $($_.Exception.Message)" 'ERROR'
    throw
  } finally {
    Stop-Spinner -Id $Id
  }
}

# ----------------------------- Helpers -----------------------------
function Ensure-Tls12 {
  [CmdletBinding()]
  param()
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { Write-Log "Failed to set TLS 1.2: $($_.Exception.Message)" 'WARN' }
}

function Ensure-ExecutionPolicy {
  [CmdletBinding()]
  param()
  try {
    $policy = Get-ExecutionPolicy -List | Where-Object { $_.Scope -eq 'Process' } | Select-Object -ExpandProperty ExecutionPolicy -ErrorAction SilentlyContinue
  } catch { $policy = $null }
  if ($null -eq $policy -or $policy -in @('Undefined','Restricted','AllSigned')) {
    try {
      Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
      Write-Log "ExecutionPolicy(Process) set to Bypass for this session." 'INFO'
    } catch {
      Write-Log "Failed to set ExecutionPolicy(Process) Bypass: $($_.Exception.Message)" 'WARN'
    }
  }
}

function Ensure-Module {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([Parameter(Mandatory)][string]$Name)
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    if ($PSCmdlet.ShouldProcess("Install-Module $Name", "Install")) {
      try { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
      Install-Module -Name $Name -Repository PSGallery -Force -Scope AllUsers -AllowClobber -SkipPublisherCheck
      Write-Log "Installed module: $Name" 'SUCCESS'
    }
  } else {
    Write-Log "Module available: $Name"
  }
}

function Test-Command {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Name)
  return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Get-RebootPending {
  [CmdletBinding()]
  param()
  try {
    if (Get-Command Get-WURebootStatus -ErrorAction SilentlyContinue) {
      $st = Get-WURebootStatus -Silent
      if ($st -and ($st.RebootRequired -or $st.IsRebootRequired)) { return $true }
    }
  } catch {}
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
    'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' # PendingFileRenameOperations
  )
  foreach ($p in $paths) { if (Test-Path $p) { return $true } }
  return $false
}

# ----------------------------- Update Blocks -----------------------------
function Update-Winget {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([switch]$EnableMSStore)
  if ($SkipWinget) { Write-Log "Skipping Winget updates."; return }
  if (-not (Test-Command -Name 'winget.exe')) {
    Write-Log "Winget not found. Install 'App Installer' from Microsoft Store, then re-run." 'WARN'
    return
  }

  try {
    if ($EnableMSStore) {
      $sources = winget source list 2>$null
      $hasMsStore = ($sources -join "`n") -match '^\s*msstore\b'
      if (-not $hasMsStore) {
        if ($PSCmdlet.ShouldProcess("winget source add msstore","Add")) {
          winget source add -n msstore -a https://storeedgefd.dsx.mp.microsoft.com/v9.0 2>$null | Out-Null
          Write-Log "Added msstore source to winget." 'SUCCESS'
        }
      } else {
        if ($PSCmdlet.ShouldProcess("winget source enable msstore","Enable")) {
          winget source enable msstore 2>$null | Out-Null
          Write-Log "Ensured msstore source enabled in winget." 'SUCCESS'
        }
      }
    }

    if ($PSCmdlet.ShouldProcess("winget source update","Update")) {
      winget source update 2>$null | Out-Null
      Write-Log "Winget sources updated." 'SUCCESS'
    }
    if ($PSCmdlet.ShouldProcess("winget upgrade --all","Upgrade all")) {
      winget upgrade --all --accept-source-agreements --accept-package-agreements --disable-interactivity 2>$null | Out-Null
      Write-Log "Winget packages upgraded (upgrade --all)." 'SUCCESS'
    }
    if ($PSCmdlet.ShouldProcess("winget update --all","Update all (alias)")) {
      winget update --all --accept-source-agreements --accept-package-agreements --disable-interactivity 2>$null | Out-Null
      Write-Log "Winget packages upgraded (update --all alias)." 'SUCCESS'
    }
  } catch {
    Write-Log "Winget update error: $($_.Exception.Message)" 'ERROR'
  }
}

function Update-Windows {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()
  if ($SkipWindowsUpdate) { Write-Log "Skipping Windows Update."; return }
  Ensure-Tls12
  Ensure-ExecutionPolicy
  Ensure-Module -Name PSWindowsUpdate

  try {
    Import-Module PSWindowsUpdate -ErrorAction Stop
    if ($PSCmdlet.ShouldProcess("Enable Microsoft Update","Add service")) {
      Add-WUServiceManager -MicrosoftUpdate -ErrorAction SilentlyContinue | Out-Null
    }
    if ($PSCmdlet.ShouldProcess("Get-WindowsUpdate -Install","Install updates")) {
      Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot | Out-Null
      Write-Log "Windows and Microsoft updates installed (reboot may be pending)." 'SUCCESS'
    }
  } catch {
    Write-Log "Windows Update error: $($_.Exception.Message)" 'ERROR'
  }
}

function Update-Defender {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()
  if ($SkipDefender) { Write-Log "Skipping Microsoft Defender tasks."; return }
  try {
    if (Test-Command -Name 'Update-MpSignature') {
      if ($PSCmdlet.ShouldProcess("Update Defender signatures","Update")) {
        Update-MpSignature | Out-Null
        Write-Log "Defender signatures updated." 'SUCCESS'
      }
    }
    if ($MinimalHardening -or $All -or (-not $UpdateOnly)) {
      if ($PSCmdlet.ShouldProcess("Set Defender preferences","Harden")) {
        try { Set-MpPreference -MAPSReporting 2 -ErrorAction Stop } catch { Set-MpPreference -MAPSReporting 1 -ErrorAction SilentlyContinue }
        try { Set-MpPreference -CloudBlockLevel 2 -ErrorAction Stop } catch { Set-MpPreference -CloudBlockLevel 1 -ErrorAction SilentlyContinue }
        try { Set-MpPreference -SubmitSamplesConsent 1 -ErrorAction Stop } catch {}
        try { Set-MpPreference -PUAProtection 1 -ErrorAction Stop } catch {}
        Write-Log "Defender minimal hardening applied (MAPS, CloudBlock, SampleConsent, PUA)." 'SUCCESS'
      }
    }
    if ($ApplyASR) {
      $ruleIds = @(
        'D4F940AB-401B-4EfC-AADC-AD5F3C50688A',
        '3B576869-A4EC-4529-8536-B80A7769E899',
        '75668C1F-73B5-4CF0-BB93-3ECF5CB7CC84',
        '5BEB7EFE-FD9A-4556-801D-275E5FFC04CC',
        'BE9BA2D9-53EA-4CDC-84E5-9B1EEEE46550',
        '26190899-1602-49E8-8B27-EB1D0A1CE869'
      )
      if ($PSCmdlet.ShouldProcess("Apply ASR rules","Enforce")) {
        Set-MpPreference -AttackSurfaceReductionRules_Ids $ruleIds -AttackSurfaceReductionRules_Actions (@(1,1,1,1,1,1))
        Write-Log "ASR rules enforced (recommended set)." 'SUCCESS'
      }
    }
    if ($PSCmdlet.ShouldProcess("Defender Quick Scan","Scan")) {
      Start-MpScan -ScanType QuickScan | Out-Null
      Write-Log "Defender quick scan started." 'SUCCESS'
    }
  } catch {
    Write-Log "Defender error: $($_.Exception.Message)" 'ERROR'
  }
}

function Repair-SystemHealth {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()
  if ($SkipHealthRepair) { Write-Log "Skipping DISM/SFC."; return }
  try {
    if ($PSCmdlet.ShouldProcess("DISM /RestoreHealth","Repair")) {
      & DISM.exe /Online /Cleanup-Image /RestoreHealth | Out-Null
      Write-Log "DISM RestoreHealth completed." 'SUCCESS'
    }
    if ($PSCmdlet.ShouldProcess("SFC /scannow","Verify/Repair")) {
      & sfc.exe /scannow | Out-Null
      Write-Log "SFC scan completed." 'SUCCESS'
    }
  } catch {
    Write-Log "Health repair error: $($_.Exception.Message)" 'ERROR'
  }
}

function Update-WSL {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()
  if ($SkipWSL) { Write-Log "Skipping WSL update."; return }
  if (Test-Command -Name 'wsl.exe') {
    try {
      if ($PSCmdlet.ShouldProcess("WSL --update","Update")) {
        & wsl.exe --update | Out-Null
        Write-Log "WSL updated (if installed)." 'SUCCESS'
      }
    } catch {
      Write-Log "WSL update error: $($_.Exception.Message)" 'WARN'
    }
  } else {
    Write-Log "WSL not installed." 'DEBUG'
  }
}

# ----------------------------- Orchestration -----------------------------
try {
  Assert-Admin
  New-LogInfra
  Banner -Text "Windows Update & Strengthen — starting"
  Write-Log "Starting update/hardening run. Args: $($PSBoundParameters.Keys -join ', ')"

  $doMinimal = $MinimalHardening -or $All -or (-not $UpdateOnly)

  Invoke-Step -Id 1 -Name "Winget: source update & app upgrades" -Hint "Downloading & upgrading apps; package installs may pause console output." -Long -Action { Update-Winget -EnableMSStore }
  Invoke-Step -Id 2 -Name "Windows Update (quality + Microsoft Update)" -Hint "Scanning and downloading patches; this phase may take time depending on bandwidth." -Long -Action { Update-Windows }
  Invoke-Step -Id 3 -Name "Microsoft Defender: update & hardening" -Action { Update-Defender }
  Invoke-Step -Id 4 -Name "WSL update" -Action { Update-WSL }
  if ($All -or -not $UpdateOnly) {
    Invoke-Step -Id 5 -Name "System health repair (DISM/SFC)" -Hint "Component store health checks; may take time on some systems." -Long -Action { Repair-SystemHealth }
  }

  $pending = Get-RebootPending
  Write-Log ("Reboot pending: {0}" -f $pending) 'INFO'

  if ($RebootIfNeeded -and $pending -and -not $NoReboot) {
    Banner -Text "Rebooting to finish updates"
    Write-Log "Rebooting in 5 seconds..." 'WARN'
    Start-Sleep -Seconds 5
    Restart-Computer -Force
  } else {
    if ($pending) {
      Banner -Text "Reboot recommended to complete updates"
      Write-Log "Reboot recommended to complete updates." 'WARN'
    }
    Banner -Text "Windows Update & Strengthen — complete"
    Write-Log "Run complete." 'SUCCESS'
  }
}
catch {
  Write-Log "Fatal: $($_.Exception.Message)" 'ERROR'
  throw
}
finally {
  Stop-LogInfra
}
