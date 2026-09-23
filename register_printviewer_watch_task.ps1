# Registers a scheduled task: run watch_printviewer.bat on an interval (default 10 min, current user).
# Usage: powershell -ExecutionPolicy Bypass -File "...\register_printviewer_watch_task.ps1" [-Force] [-IntervalMinutes 10]
# Japanese GUI steps: see タスクスケジューラの設定手順.md

param(
    [string]$TaskName = "printviewer-watch",
    [ValidateRange(1, 1439)]
    [int]$IntervalMinutes = 10,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$watchBat = Join-Path $root "watch_printviewer.bat"
if (-not (Test-Path -LiteralPath $watchBat)) {
    Write-Error "watch_printviewer.bat not found: $watchBat"
}

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    if (-not $Force) {
        Write-Host "Task '$TaskName' already exists. Re-run with -Force to replace."
        exit 1
    }
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$action = New-ScheduledTaskAction -Execute $watchBat -WorkingDirectory $root
$start = (Get-Date).AddSeconds(-30)
$triggerRepeat = New-ScheduledTaskTrigger -Once -At $start `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
$triggerLogon = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$triggers = @($triggerRepeat, $triggerLogon)
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers `
    -Settings $settings -Principal $principal `
    -Description "Restart printviewer if python is not running; every $IntervalMinutes min + at logon (watch_printviewer.bat)"

Write-Host "Registered scheduled task: $TaskName (every $IntervalMinutes min + at logon)"
Write-Host "Open Task Scheduler (taskschd.msc) or test: Start-ScheduledTask -TaskName '$TaskName'"
