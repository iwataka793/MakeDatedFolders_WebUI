#requires -Version 5.1
[CmdletBinding()]
param(
  [int]$Port = 0,
  [switch]$NoBrowser,
  [switch]$AppMode
)

$ErrorActionPreference = 'Stop'

function Minimize-ConsoleWindow {
  try {
    $sig = '[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
    Add-Type -Namespace Win32 -Name User32 -MemberDefinition $sig -ErrorAction SilentlyContinue | Out-Null
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    if ($hwnd -ne 0) { [Win32.User32]::ShowWindow($hwnd, 6) | Out-Null }
  } catch {}
}

# --- Project root (robust for ISE/console) ---
$root = $PSScriptRoot
if (-not $root -or $root.Trim() -eq '') {
  if ($PSCommandPath) { $root = Split-Path -Parent $PSCommandPath }
}
if (-not $root -or $root.Trim() -eq '') { $root = (Get-Location).Path }

Set-Location $root

# --- Best-effort: Unblock files if the ZIP had Mark-of-the-Web ---
$unblock = Get-Command Unblock-File -ErrorAction SilentlyContinue
if ($unblock) {
  Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -in @('.ps1','.psm1','.js','.css','.html') } |
    ForEach-Object {
      try { Unblock-File -LiteralPath $_.FullName -ErrorAction Stop } catch {}
    }
}

# --- Ensure STA (needed for FolderBrowserDialog) ---
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $self = $PSCommandPath
  $args = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File', $self)
  if ($Port -gt 0) { $args += @('-Port', $Port) }
  if ($NoBrowser) { $args += @('-NoBrowser') }
  if ($AppMode) { $args += @('-AppMode') }
  $windowStyle = if ($AppMode) { 'Minimized' } else { 'Normal' }
  Start-Process -FilePath 'powershell.exe' -WindowStyle $windowStyle -ArgumentList $args | Out-Null
  return
}

# --- Try to relax policy for this process (if allowed) ---
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force } catch {}
if ($AppMode) { Minimize-ConsoleWindow }

$server = Join-Path $root 'host\FolderMaker.Server.ps1'
if (-not (Test-Path -LiteralPath $server)) {
  throw "Server script not found: $server"
}

& $server -Port $Port -HostName 'localhost' -NoOpenBrowser:$NoBrowser -AppMode:$AppMode
