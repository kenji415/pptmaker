# Create desktop shortcut to test_scan_printer.bat
$base = $PSScriptRoot
if (-not $base) { $base = Split-Path -Parent $MyInvocation.MyCommand.Path }

$WshShell = New-Object -ComObject WScript.Shell
$Desktop = [Environment]::GetFolderPath('Desktop')
$Shortcut = $WshShell.CreateShortcut("$Desktop\test_scan_printer.lnk")
$Shortcut.TargetPath = Join-Path $base "test_scan_printer.bat"
$Shortcut.WorkingDirectory = $base
$Shortcut.Description = "test_scan_printer"
$Shortcut.Save()

Write-Host "Desktop shortcut created: test_scan_printer.lnk"
