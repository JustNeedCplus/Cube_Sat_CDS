# Build CubeSat Builder shortcut for Windows
# Run this in PowerShell from the project directory
# Source code edits take effect immediately - no rebuild needed

$SrcDir = $PSScriptRoot
if (-not $SrcDir) { $SrcDir = Get-Location }

$AppName = "CubeSat Builder"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = "$DesktopPath\$AppName.lnk"

# Create a batch launcher
$LauncherPath = "$SrcDir\launch.bat"
@"
@echo off
cd /d "$SrcDir"
call .venv\Scripts\activate
python main.py %*
"@ | Out-File -FilePath $LauncherPath -Encoding ASCII

Write-Host ""
Write-Host "Created launcher: $LauncherPath"
Write-Host ""

# Create desktop shortcut
try {
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $LauncherPath
    $Shortcut.WorkingDirectory = $SrcDir
    $Shortcut.Description = "CubeSat Builder - Subsystem Sizing & BOM Tool"
    $Shortcut.WindowStyle = 7  # Minimized (hides the cmd window faster)
    $Shortcut.Save()
    Write-Host "Desktop shortcut created: $ShortcutPath"
} catch {
    Write-Host "Could not create shortcut automatically."
    Write-Host "You can manually create a shortcut to: $LauncherPath"
}

Write-Host ""
Write-Host "IMPORTANT: Source code lives at $SrcDir"
Write-Host "Edit main.py or CubeSat-Builder.html and changes take effect immediately."
Write-Host "Only re-run this script if you move the source folder."
Write-Host ""
Write-Host "Double-click the desktop shortcut or run launch.bat to start the app."
