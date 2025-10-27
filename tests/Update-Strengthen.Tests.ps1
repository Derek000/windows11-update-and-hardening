# Pester tests for presence of progress helpers and logging
Describe "Update-Strengthen functions" {
  It "should define progress helpers" {
    . "$PSScriptRoot\..\Update-Strengthen.ps1" -WhatIf
    Get-Command Invoke-Step -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    Get-Command Start-Spinner -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    Get-Command Stop-Spinner -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
  }
}

Describe "Update-Strengthen logging (no-op run)" {
  $scriptPath = Join-Path $PSScriptRoot "..\Update-Strengthen.ps1"
  $logRoot = Join-Path $env:ProgramData "Madrock\Update-Strengthen\logs"

  It "runs with skip flags and creates logs" {
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    $beforeOps = Get-ChildItem $logRoot -Filter "*_ops.log" -ErrorAction SilentlyContinue
    $beforeTrn = Get-ChildItem $logRoot -Filter "*_transcript.log" -ErrorAction SilentlyContinue

    & pwsh -NoProfile -File $scriptPath -UpdateOnly -SkipWinget -SkipWindowsUpdate -SkipDefender -SkipHealthRepair -SkipWSL -NoProgress

    $afterOps = Get-ChildItem $logRoot -Filter "*_ops.log" -ErrorAction SilentlyContinue
    $afterTrn = Get-ChildItem $logRoot -Filter "*_transcript.log" -ErrorAction SilentlyContinue

    ($afterOps.Count -gt $beforeOps.Count) | Should -BeTrue
    ($afterTrn.Count -gt $beforeTrn.Count) | Should -BeTrue
  }
}
