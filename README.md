# Windows 11 — Update & Strengthen (One Command)

## Why
Timely patching and a sane baseline hardening cut risk from known CVEs and commodity malware. This script updates Windows & apps and applies conservative Defender hardening, with optional stronger controls.

## What it does
- **Windows Update** (quality & Microsoft Update) via `PSWindowsUpdate`
- **Winget** sources + upgrades for installed apps (incl. MS Store source)
- **Microsoft Defender**: signature update, quick scan, conservative hardening (PUA block, cloud protection, MAPS, safe sample submission)
- **System health**: `DISM /RestoreHealth`, `SFC /scannow`
- **WSL update** (if installed)
- **Interactive progress**: console banners + spinner-based `Write-Progress` so you can see where time is being spent (disable via `-NoProgress`).
- **Logging**: 
  - **Ops log** `C:\ProgramData\Madrock\Update-Strengthen\logs\run_YYYYMMDD_HHMMSS_ops.log`
  - **Transcript** `C:\ProgramData\Madrock\Update-Strengthen\logs\run_YYYYMMDD_HHMMSS_transcript.log`
  - Windows Event Log source: `Madrock.UpdateStrengthen`

## Requirements
- Windows 11, **run as Administrator**
- Winget (App Installer) recommended
- PowerShell 5.1+ or PowerShell 7.x
- Internet access to Windows Update, PSGallery, Winget sources

> If you see `PSWindowsUpdate.psm1 cannot be loaded because running scripts is disabled`, the script sets **ExecutionPolicy(Process)=Bypass** automatically for this session.

## Quick start
```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\Update-Strengthen.ps1 -All -RebootIfNeeded
```

### Options (common)
- Updates only:
  ```powershell
  .\Update-Strengthen.ps1 -UpdateOnly
  ```
- Update + conservative hardening (default behaviour):
  ```powershell
  .\Update-Strengthen.ps1
  ```
- Stronger hardening with ASR rules (may impact Office/macros/dev tooling):
  ```powershell
  .\Update-Strengthen.ps1 -All -ApplyASR
  ```
- Suppress progress UI:
  ```powershell
  .\Update-Strengthen.ps1 -NoProgress
  ```

## Winget MS Store source
If `msstore` source is missing, the script attempts to add it:
```
winget source add -n msstore -a https://storeedgefd.dsx.mp.microsoft.com/v9.0
```
Some environments may restrict MS Store access; updates still proceed for other sources.

## Security alignment
- **OWASP**: secure configuration & patch management
- **NIST CSF**: PR.IP-12, PR.IP-1, DE.CM-4
- **CIS Controls**: 1, 3, 5, 7, 8, 12
- **CIS Benchmarks (Windows 11)**: conservative items (PUA, cloud protection). ASR rules optional.

## Uninstall / revert
- Defender preferences can be reset in Windows Security → Virus & threat protection → Manage settings.
- ASR rules can be turned off via:
  ```powershell
  Set-MpPreference -AttackSurfaceReductionRules_Ids @('D4F940AB-401B...','...') -AttackSurfaceReductionRules_Actions @(0,0,0,0,0,0)
  ```
