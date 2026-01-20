#requires -Version 5.1
[CmdletBinding()]
param(
  [int]$Port = 0,
  [string]$HostName = 'localhost',
  [switch]$NoOpenBrowser,
  [switch]$AppMode
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $root 'core\FolderMaker.Core.psm1'
# Exported function名に "未承認の動詞" が含まれていても動作には問題ありませんが、毎回の警告はうるさいので抑止します。
$__wp = $WarningPreference
try {
  $WarningPreference = 'SilentlyContinue'
  Import-Module $modulePath -Force -DisableNameChecking
} finally {
  $WarningPreference = $__wp
}
Initialize-FolderMakerConfig
function Get-ContentBytes([string]$Text){ [Text.Encoding]::UTF8.GetBytes($Text) }
function Get-RequestEncoding([System.Net.HttpListenerRequest]$req){
  # HttpListener defaults ContentEncoding to system ANSI when charset is omitted.
  # Our WebUI sends UTF-8 JSON, so default to UTF-8 unless charset is explicitly provided.
  try {
    if($req.ContentType -and $req.ContentType -match "charset\s*=\s*([^;]+)"){
      return [Text.Encoding]::GetEncoding($matches[1].Trim())
    }
  } catch {}
  return [Text.Encoding]::UTF8
}
function Read-RequestBody([System.Net.HttpListenerRequest]$req){
  $enc = Get-RequestEncoding $req
  $ms = New-Object IO.MemoryStream
  try {
    $req.InputStream.CopyTo($ms)
    $bytes = $ms.ToArray()
    return $enc.GetString($bytes)
  } finally {
    $ms.Dispose()
  }
}
function Write-Bytes([System.Net.HttpListenerResponse]$res,[byte[]]$bytes,[string]$contentType,[int]$statusCode=200){ $res.StatusCode=$statusCode; $res.ContentType=$contentType; $res.ContentLength64=$bytes.Length; $res.OutputStream.Write($bytes,0,$bytes.Length); $res.Close() }
function Write-Json([System.Net.HttpListenerResponse]$res,$obj,[int]$statusCode=200){ $json = $obj | ConvertTo-Json -Depth 12 -Compress; Write-Bytes $res (Get-ContentBytes $json) 'application/json; charset=utf-8' $statusCode }
function Mime-FromExt([string]$path){ switch ([IO.Path]::GetExtension($path).ToLowerInvariant()) { '.html' { 'text/html; charset=utf-8' } '.css' { 'text/css; charset=utf-8' } '.js' { 'text/javascript; charset=utf-8' } '.svg' { 'image/svg+xml' } '.png' { 'image/png' } '.jpg' { 'image/jpeg' } '.jpeg' { 'image/jpeg' } default { 'application/octet-stream' } } }
function Try-StartListener([int]$p){
  $listener = New-Object System.Net.HttpListener
  $prefix = "http://$HostName`:$p/"
  $listener.Prefixes.Add($prefix) | Out-Null
  try { $listener.Start(); return @($true,$listener,$prefix) } catch { try { $listener.Stop() } catch {} ; try { $listener.Close() } catch {} ; return @($false,$null,$prefix) }
}
function Resolve-BrowserPath([string]$exeName,[string[]]$fallbacks){
  $cmd = Get-Command $exeName -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  foreach ($candidate in $fallbacks) {
    if ($candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }
  }
  return $null
}
function Convert-PlanItem($item){
  $dateStr = ''
  try {
    if ($item.Date) { $dateStr = ([datetime]$item.Date).ToString('yyyy-MM-dd') }
  } catch {
    $dateStr = ''
  }
  return [pscustomobject]@{
    Kind       = $item.Kind
    Date       = $dateStr
    Index      = $item.Index
    FolderName = $item.FolderName
    FullPath   = $item.FullPath
    Exists     = $item.Exists
    Action     = $item.Action
  }
}
function Convert-RunItem($item){
  $result = [string]$item.Result
  $kind = if ($result -match 'Year') { 'Year' } elseif ($result -match 'Month') { 'Month' } else { 'Day' }
  $action = if ($result -match '^Created') { 'Create' } elseif ($result -match '^Skipped') { 'Skip' } else { '' }
  return [pscustomobject]@{
    Kind       = $kind
    Date       = ''
    FolderName = $item.FolderName
    Action     = $action
    FullPath   = $item.FullPath
  }
}

if ($Port -le 0) {
  $portsToTry = 8787..8797
} else {
  $portsToTry = @($Port)
}
$listener = $null
$prefix = $null
foreach ($p in $portsToTry) {
  $r = Try-StartListener -p $p
  if ($r[0]) { $listener = $r[1]; $prefix = $r[2]; break }
}
if (-not $listener) { throw 'ポートを開けませんでした。別のポートを指定してください。' }
$uiRoot = Join-Path $root 'ui'
$script:lastPing = Get-Date
$script:lastPingLogAt = $null
$script:listenerRef = $listener
$script:closeRequestedAt = $null
$script:shutdownRequested = $false
$script:ActiveRequests = 0
$script:timerRef = $null
$timer = $null
if (-not $NoOpenBrowser) {
  $script:AutoShutdownSeconds = 120
  $script:CloseGraceSeconds = 30
  $timer = New-Object System.Timers.Timer
  $timer.Interval = 3000
  $timer.AutoReset = $true
  $timer.add_Elapsed({
    if ($script:ActiveRequests -gt 0) {
      return
    }
    $closeAt = $script:closeRequestedAt
    if ($null -ne $closeAt) {
      $closeDelta = (New-TimeSpan -Start $closeAt -End (Get-Date)).TotalSeconds
      if ($closeDelta -ge $script:CloseGraceSeconds) {
        Write-Host '[FolderMaker] close requested -> stopping server'
        try { $script:listenerRef.Stop() } catch {}
        try { $script:listenerRef.Close() } catch {}
        try { $this.Stop() } catch {}
      }
      return
    }
    $last = $script:lastPing
    if ($null -ne $last) {
      $delta = (New-TimeSpan -Start $last -End (Get-Date)).TotalSeconds
      if ($delta -ge $script:AutoShutdownSeconds) {
        Write-Host ('[FolderMaker] auto-shutdown: idle {0:N1}s' -f $delta)
        try { $script:listenerRef.Stop() } catch {}
        try { $script:listenerRef.Close() } catch {}
        try { $this.Stop() } catch {}
      }
    }
  })
  $timer.Start()
  $script:timerRef = $timer
}
Write-Host ("[FolderMaker] listening: {0}" -f $prefix)
Write-Host '[FolderMaker] stop: Ctrl+C'
if (-not $NoOpenBrowser) {
  try {
    if ($AppMode) {
      # Chrome/Edge の "アプリモード" で開ければ、アドレスバー無しのウィンドウになって見た目がアプリっぽくなります。
      # NOTE: Windows PowerShell 5.1 では null 条件演算子 (?.) が使えません。
      $edge = Resolve-BrowserPath 'msedge.exe' @(
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:LocalAppData 'Microsoft\Edge\Application\msedge.exe')
      )
      $chrome = Resolve-BrowserPath 'chrome.exe' @(
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'),
        (Join-Path $env:LocalAppData 'Google\Chrome\Application\chrome.exe')
      )
      if ($edge) {
        Start-Process -FilePath $edge -ArgumentList @("--app=$prefix")
      } elseif ($chrome) {
        Start-Process -FilePath $chrome -ArgumentList @("--app=$prefix")
      } else {
        Start-Process $prefix
      }
    } else {
      Start-Process $prefix
    }
  } catch {}
}
while ($listener.IsListening) {
  $ctx = $null
  try { $ctx = $listener.GetContext() } catch { break }
  $req = $ctx.Request
  $res = $ctx.Response
  $script:lastPing = Get-Date
  try {
    $path = $req.Url.AbsolutePath
    if ($path -eq '/favicon.ico') {
      Write-Bytes $res ([byte[]]@()) 'image/x-icon' 204
      continue
    }
    if ($path -eq '/api/health' -and $req.HttpMethod -eq 'GET') { Write-Json $res @{ok=$true;ts=(Get-Date).ToString('s')} ; continue }
    if ($path -eq '/api/ping' -and $req.HttpMethod -eq 'POST') {
      $script:lastPing = Get-Date
      if ($null -eq $script:lastPingLogAt -or (New-TimeSpan -Start $script:lastPingLogAt -End $script:lastPing).TotalSeconds -ge 30) {
        Write-Host '[FolderMaker] ping'
        $script:lastPingLogAt = $script:lastPing
      }
      Write-Json $res @{ ok = $true } 200
      continue
    }
    if ($path -eq '/api/close' -and $req.HttpMethod -eq 'POST') {
      # UI ウィンドウが閉じられたらサーバも終了 → Start.bat の PowerShell ホストも閉じる
      $script:closeRequestedAt = Get-Date
      Write-Host '[FolderMaker] close requested'
      Write-Json $res @{ ok = $true } 200
      continue
    }
    if ($path -eq '/api/shutdown' -and $req.HttpMethod -eq 'POST') {
      $script:closeRequestedAt = Get-Date
      $script:shutdownRequested = $true
      Write-Host '[FolderMaker] shutdown requested'
      Write-Json $res @{ ok = $true } 200
      if (-not $script:timerRef) {
        try { $script:listenerRef.Stop() } catch {}
        try { $script:listenerRef.Close() } catch {}
      }
      continue
    }
    if ($path -eq '/api/config/basePath' -and $req.HttpMethod -eq 'POST') {
      $bodyText = Read-RequestBody $req
      $payload = $null
      try { $payload = $bodyText | ConvertFrom-Json } catch { $payload = $null }
      $basePath = if ($payload -and $payload.basePath) { [string]$payload.basePath } else { '' }
      if (-not $basePath) { Write-Json $res @{ ok = $false; errors = @('作成先パスが空です') } 400; continue }
      if (-not (Test-Path -LiteralPath $basePath)) { Write-Json $res @{ ok = $false; errors = @(('作成先パスが存在しません: {0}' -f $basePath)) } 400; continue }
      $cfgPath = Get-FolderMakerConfigPath
      $cfg = Load-FolderMakerConfig -Path $cfgPath
      $cfg['DefaultBasePath'] = $basePath
      Save-FolderMakerConfig -Path $cfgPath -Config $cfg
      Write-Json $res @{ ok = $true; configPath = $cfgPath } 200
      continue
    }
    if ($path -eq '/api/config' -and $req.HttpMethod -eq 'GET') {
      $cfgPath = Get-FolderMakerConfigPath
      $cfg = Load-FolderMakerConfig -Path $cfgPath
      Write-Json $res @{ ok=$true; configPath=$cfgPath; config=$cfg }
      continue
    }
    if ($path -eq '/api/pickFolder' -and $req.HttpMethod -eq 'POST') {
      # Windows側のダイアログでフォルダを選ぶ（ローカルのみ）
      # FolderBrowserDialog より見た目が現代的になりやすい OpenFileDialog を「フォルダ選択風」に使用
      $bodyText = Read-RequestBody $req
      $payload = $null
      try { $payload = $bodyText | ConvertFrom-Json } catch { $payload = $null }
      $initial = if ($payload -and $payload.initialPath) { [string]$payload.initialPath } else { '' }
      try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null

        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = '作成先フォルダを選択してください'
        $dlg.CheckFileExists = $false
        $dlg.CheckPathExists = $true
        $dlg.ValidateNames = $false
        $dlg.DereferenceLinks = $true
        $dlg.FileName = 'フォルダを選択'
        if ($initial -and (Test-Path -LiteralPath $initial)) { $dlg.InitialDirectory = $initial }

        $owner = New-Object System.Windows.Forms.Form
        $owner.ShowInTaskbar = $false
        $owner.WindowState = 'Minimized'
        $owner.TopMost = $true
        $owner.Opacity = 0
        $owner.StartPosition = 'CenterScreen'
        $owner.Show()
        $owner.Activate()

        try {
          $r = $dlg.ShowDialog($owner)
        } finally {
          $owner.Close()
          $owner.Dispose()
        }
        if ($r -ne [System.Windows.Forms.DialogResult]::OK) {
          Write-Json $res @{ ok = $false; canceled = $true } 200
          continue
        }

        $picked = [string]$dlg.FileName
        $p = Split-Path -Path $picked -Parent
        if (-not $p -or $p.Trim() -eq '') { $p = $picked }

        Write-Json $res @{ ok = $true; path = $p } 200
        continue
      } catch {
        Write-Json $res @{ ok = $false; errors = @($_.Exception.Message) } 500
        continue
      }
    }
    if (($path -eq '/api/preview' -or $path -eq '/api/run') -and $req.HttpMethod -eq 'POST') {
      $opName = if ($path -eq '/api/preview') { 'preview' } else { 'run' }
      $script:ActiveRequests += 1
      $sw = [Diagnostics.Stopwatch]::StartNew()
      try {
        $bodyText = Read-RequestBody $req
        $payload = $null
        try { $payload = $bodyText | ConvertFrom-Json } catch { Write-Json $res @{ok=$false;errors=@('JSONが不正です');raw=$bodyText} 400; continue }
        $basePath = [string]$payload.basePath
        $mode = [string]$payload.mode
        $foldersPerDay = [int]$payload.foldersPerDay
        $firstDayStartIndex = [int]$payload.firstDayStartIndex
        $startDate = [datetime]$payload.startDate
        $endDate = $null
        if ($mode -eq 'Range') { $endDate = [datetime]$payload.endDate } else { $daysToMake = [int]$payload.daysToMake; $endDate = $startDate.Date.AddDays($daysToMake-1) }
        $errors = New-Object System.Collections.Generic.List[string]
        if (-not $basePath) { $errors.Add('作成先パスが空です') }
        if ($foldersPerDay -lt 1) { $errors.Add('最大番号は1以上にしてください') }
        if ($firstDayStartIndex -lt 1) { $errors.Add('初日の開始番号は1以上にしてください') }
        if ($mode -ne 'Range' -and $mode -ne 'Days') { $errors.Add('mode は Range / Days のどちらかです') }
        if ($mode -eq 'Days' -and $daysToMake -lt 1) { $errors.Add('日数は1以上にしてください') }
        if ($endDate.Date -lt $startDate.Date) { $errors.Add('終了日は開始日以降にしてください') }
        if (-not (Test-Path -LiteralPath $basePath)) { $errors.Add(('作成先パスが存在しません: {0}' -f $basePath)) }
        if ($errors.Count -gt 0) { Write-Json $res @{ok=$false;errors=$errors} 400; continue }
        if ($path -eq '/api/preview') {
          Write-Host ('[FolderMaker] preview start: {0}' -f $basePath)
          $plan = Get-PlannedFolderList -BasePath $basePath -StartDate $startDate -EndDate $endDate -FoldersPerDay $foldersPerDay -FirstDayStartIndex $firstDayStartIndex
          $days = ($endDate.Date - $startDate.Date).Days + 1
	        # NOTE: PowerShell は 1件だけヒットすると配列ではなくスカラーになるため、@() で必ず配列化して .Count を安全に取る
	        $summary = @{
	          days  = $days
	          total = @($plan).Count
	          create = @($plan | Where-Object { $_.Action -eq 'Create' }).Count
	          skip   = @($plan | Where-Object { $_.Action -eq 'Skip'   }).Count
	          start = $startDate.ToString('yyyy-MM-dd')
	          end   = $endDate.ToString('yyyy-MM-dd')
	        }
          $items = foreach ($item in @($plan)) { Convert-PlanItem $item }
          Write-Json $res @{ok=$true; summary=$summary; items=$items }
          continue
        } else {
          Write-Host ('[FolderMaker] run start: {0}' -f $basePath)
          $result = New-DateIndexedFolders -BasePath $basePath -StartDate $startDate -EndDate $endDate -FoldersPerDay $foldersPerDay -FirstDayStartIndex $firstDayStartIndex -Confirm:$false
          $items = foreach ($item in @($result)) { Convert-RunItem $item }
	        Write-Json $res @{
	          ok = $true
	          created = @($result | Where-Object { $_.Result -match '^Created' }).Count
	          skipped = @($result | Where-Object { $_.Result -match '^Skipped' }).Count
	          items = $items
	        }
          continue
        }
      } finally {
        $sw.Stop()
        $script:ActiveRequests = [Math]::Max(0, ($script:ActiveRequests - 1))
        $script:lastPing = Get-Date
        Write-Host ('[FolderMaker] {0} end ({1:N1}s, active={2})' -f $opName, $sw.Elapsed.TotalSeconds, $script:ActiveRequests)
      }
    }
    # static files
    if ($path -eq '/') { $path = '/index.html' }
    $safePath = $path.TrimStart('/').Replace('/', [IO.Path]::DirectorySeparatorChar)
    $full = Join-Path $uiRoot $safePath
    if (-not ($full.StartsWith($uiRoot,[StringComparison]::OrdinalIgnoreCase))) { Write-Bytes $res (Get-ContentBytes 'Bad Request') 'text/plain; charset=utf-8' 400; continue }
    if (-not (Test-Path -LiteralPath $full)) { Write-Bytes $res (Get-ContentBytes 'Not Found') 'text/plain; charset=utf-8' 404; continue }
    $bytes = [IO.File]::ReadAllBytes($full)
    Write-Bytes $res $bytes (Mime-FromExt $full) 200
  } catch {
    try { Write-Json $res @{ok=$false;errors=@($_.Exception.Message)} 500 } catch {}
  }
}
try { $listener.Stop() } catch {}
try { $listener.Close() } catch {}
if ($timer) {
  try { $timer.Stop() } catch {}
  try { $timer.Dispose() } catch {}
}
