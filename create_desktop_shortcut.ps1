$WshShell = New-Object -ComObject WScript.Shell
$Desktop = [Environment]::GetFolderPath('Desktop')
$Shortcut = $WshShell.CreateShortcut("$Desktop\Scan Router.lnk")
$Shortcut.TargetPath = "C:\Users\doctor\printviewer\start_scan_router.bat"
$Shortcut.WorkingDirectory = "C:\Users\doctor\printviewer"
$Shortcut.Description = "Scan Router"
$Shortcut.Save()
Write-Host "Desktop shortcut created: Scan Router.lnk"

















