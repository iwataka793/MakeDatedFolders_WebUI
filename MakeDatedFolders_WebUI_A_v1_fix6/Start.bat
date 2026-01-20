@echo off
setlocal
cd /d "%~dp0"
echo Starting FolderMaker WebUI...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "try { Get-ChildItem -LiteralPath '%~dp0' -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.ps1','.psm1','.js','.css','.html' } | ForEach-Object { try { Unblock-File -LiteralPath $_.FullName -ErrorAction Stop } catch {} } } catch {}"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Minimized -File "%~dp0Start-WebUI.ps1" -AppMode
if errorlevel 1 (
  echo.
  echo Failed to start FolderMaker WebUI.
)
