# UX: interactive progress/spinner banners + earlier fixes

## Why
It wasn’t obvious that long-running steps (winget/Windows Update/DISM) were doing work.

## What
- **Interactive progress UI**:
  - `Invoke-Step` wrapper prints a banner and runs a **spinner-based Write-Progress** in the current runspace.
  - Clear hints for long phases (e.g., “downloading & upgrading apps; package installs may pause console output.”).
  - Disable with `-NoProgress`.
- Retains prior fixes:
  - Advanced functions for `ShouldProcess` (no `$PSCmdlet` nulls)
  - ExecutionPolicy(Process)=Bypass at runtime
  - Robust `msstore` handling
  - Split logs (`*_ops.log` & `*_transcript.log`) + lock-safe writer

## Risk
Low. UI-only when interactive; use `-NoProgress` for scripts/CI.

## Test
- Pester asserts presence of progress helpers and logging files in a no-op run.

## Screenshot (example)
Console shows:
```
==============================================================================
Windows Update & Strengthen — starting
==============================================================================
[11:22:08] Winget: source update & app upgrades
  > Downloading & upgrading apps; package installs may pause console output.
[progress bar updates here]
...
```
