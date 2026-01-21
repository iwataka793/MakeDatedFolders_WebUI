#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
function Convert-ToHalfWidthDigits {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Text)

    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int][char]$ch
        # 全角 '０'(FF10)〜'９'(FF19)
        if ($code -ge 0xFF10 -and $code -le 0xFF19) {
            [void]$sb.Append([char]($code - 0xFF10 + 0x30))
        } else {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString()
}

# --- 設定ファイル（ユーザーが編集できる） ---
function Get-FolderMakerConfigPath {
    [CmdletBinding()]
    param()
    if ($env:FOLDERMAKER_CONFIG) { return $env:FOLDERMAKER_CONFIG }
    $root = Split-Path -Parent $PSScriptRoot
    if (-not $root) { $root = (Get-Location).Path }
    return (Join-Path $root 'FolderMaker.config.ini')
}

function Get-MonthFromName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $n = Convert-ToHalfWidthDigits -Text $Name
    if ($n -match '(?<!\d)(?<m>1[0-2]|0?[1-9])(?=月)') { return [int]$Matches['m'] }
    if ($n -match '^(?<m>1[0-2]|0?[1-9])$') { return [int]$Matches['m'] }
    return $null
}

function Get-YearFromName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $n = Convert-ToHalfWidthDigits -Text $Name
    if ($n -match '(?<!\d)(?<y>\d{4})(?!\d)') { return [int]$Matches['y'] }
    return $null
}

function Get-BasePathInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$BasePath)

    $leaf = Split-Path -Leaf $BasePath
    $month = $null
    $year = $null
    if ($leaf) {
        $month = Get-MonthFromName -Name $leaf
        $year = Get-YearFromName -Name $leaf
    }

    $kind = 'Root'
    if ($null -ne $month) { $kind = 'Month' }
    elseif ($null -ne $year) { $kind = 'Year' }

    return [pscustomobject]@{
        Kind  = $kind
        Year  = $year
        Month = $month
        Leaf  = $leaf
    }
}

function Convert-ToBoolOrDefault {
    param([object]$Value, [bool]$Default)
    if ($null -eq $Value) { return $Default }
    $s = [string]$Value
    if ($s -match '^(1|true|yes|y|on)$') { return $true }
    if ($s -match '^(0|false|no|n|off)$') { return $false }
    return $Default
}

function Load-FolderMakerConfig {
    [CmdletBinding()]
    param([string]$Path)

    $cfg = @{}
    if (-not (Test-Path -LiteralPath $Path)) { return $cfg }

    foreach ($line in (Get-Content -LiteralPath $Path -Encoding Unicode -ErrorAction SilentlyContinue)) {
        $t = $line.Trim()
        if (-not $t) { continue }
        if ($t.StartsWith(';') -or $t.StartsWith('#')) { continue }
        if ($t -match '^\[.*\]$') { continue } # section
        $parts = $t.Split('=', 2)
        if ($parts.Count -ne 2) { continue }
        $k = $parts[0].Trim()
        $v = $parts[1].Trim()
        if ($k) { $cfg[$k] = $v }
    }
    return $cfg
}

function Save-FolderMakerConfig {
    [CmdletBinding()]
    param(
        [string]$Path,
        [hashtable]$Config
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('; FolderMaker.config.ini')
    $lines.Add('; 形式: Key=Value（ユーザー編集OK）')
    $lines.Add('')
    foreach ($k in ($Config.Keys | Sort-Object)) {
        $lines.Add(('{0}={1}' -f $k, $Config[$k]))
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $lines | Set-Content -LiteralPath $Path -Encoding Unicode
}

function Initialize-FolderMakerConfig {
    $path = Get-FolderMakerConfigPath
    $cfg = Load-FolderMakerConfig -Path $path

    # 既定値（ファイルが無ければ自動生成）
    $defaults = @{
        DefaultBasePath        = 'M:\008_茨城工場_製造課\第2製造部\共有\SP共有\■■第2製造部　日報■■'
        UseFiscalYear          = 'false'   # true にすると年度開始月で年度を計算
        FiscalYearStartMonth   = '1'       # 1=1月開始, 4=4月開始 など
        YearSuffix             = '年度'    # '年' / '年度' など
        MonthZeroPad           = 'false'   # true=01月, false=1月
        AllowCombinedYearMonth = 'true'    # 例: 2026年1月 / 2026年度1月 を「月フォルダ」として認識
        SearchDepthForMMdd     = '3'       # 0=Base直下のみ, 3=Base→年→月→日(MMdd-n) まで探索
    }

    $needSave = $false
    foreach ($k in $defaults.Keys) {
        if (-not $cfg.ContainsKey($k)) { $cfg[$k] = $defaults[$k]; $needSave = $true }
    }
    if ($needSave) { Save-FolderMakerConfig -Path $path -Config $cfg }

    $script:FolderMakerConfig = $cfg
}

# --- 年度/結合フォルダ/MMdd-n を吸収するコアロジック（A/B案統合） ---
function Get-FiscalYear {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Date,
        [Parameter(Mandatory)][ValidateRange(1,12)][int]$StartMonth
    )
    if ($Date.Month -ge $StartMonth) { return $Date.Year }
    return ($Date.Year - 1)
}

function Find-ExistingMonthContainerByMMdd {
    <#
      A案: ルート配下で "MMdd-n" を探索し、見つかった場合は「その親フォルダ」を返す。
      - 深さ制限つき（PS5.1 は -Depth が無いので BFS）
      - 候補が複数あれば「月っぽい/年っぽい」場所を優先
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$DateString, # MMdd
        [Parameter(Mandatory)][ValidateRange(0,50)][int]$MaxDepth,
        [int]$Year,
        [int]$Month
    )

    if (-not (Test-Path -LiteralPath $BasePath)) { return $null }

    $pattern = ('^{0}-\d+$' -f [Regex]::Escape($DateString))

    $queue = New-Object System.Collections.Generic.Queue[object]
    $queue.Enqueue(@($BasePath, 0))

    $candidates = New-Object System.Collections.Generic.List[object]

    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()
        $path = [string]$node[0]
        $depth = [int]$node[1]

        if ($depth -gt $MaxDepth) { continue }

        $children = $null
        try { $children = Get-ChildItem -LiteralPath $path -Directory -ErrorAction Stop } catch { $children = @() }

        foreach ($d in $children) {
            $nameNorm = Convert-ToHalfWidthDigits -Text $d.Name

            if ($nameNorm -match $pattern) {
                $parent = Split-Path -Parent $d.FullName
                # Base配下安全チェック
                if (-not ($parent.StartsWith($BasePath, [StringComparison]::OrdinalIgnoreCase))) { continue }

                # スコアリング
                $score = 0
                if ($null -ne $Month) {
                    if ($parent -match ('(^|\\)0?{0}月($|\\)' -f $Month)) { $score += 50 }
                }
                if ($null -ne $Year) {
                    if ($parent -match ('(^|\\){0}(\D|$)' -f $Year)) { $score += 30 }
                }
                if ($parent -match '年度') { $score += 5 }
                if ($parent -match '月') { $score += 5 }

                $candidates.Add([pscustomobject]@{
                    Parent = $parent
                    Score  = $score
                    Depth  = $depth
                })
                continue
            }

            if ($depth -lt $MaxDepth) {
                $queue.Enqueue(@($d.FullName, ($depth + 1)))
            }
        }
    }

    if ($candidates.Count -eq 0) { return $null }

    $best = $candidates | Sort-Object @{Expression='Score';Descending=$true}, @{Expression='Depth';Descending=$false}, @{Expression='Parent';Descending=$false} | Select-Object -First 1
    return $best.Parent
}

function Find-CombinedYearMonthFolder {
    <#
      B案: "2026年1月" / "2026年度1月" のような「結合フォルダ」を月フォルダとして認識する
      - 対象: BasePath 直下のフォルダ
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][int]$Year,
        [Parameter(Mandatory)][int]$Month
    )

    if (-not (Test-Path -LiteralPath $BasePath)) { return $null }

    $y = [string]$Year
    $m1 = [string]$Month
    $m2 = ('{0:00}' -f $Month)

    $rxYear = ('(?<!\d){0}(?!\d)' -f [Regex]::Escape($y))
    $rxMonth = ('(?<!\d)({0}|{1})(?!\d)' -f [Regex]::Escape($m1), [Regex]::Escape($m2))

    try { $dirs = Get-ChildItem -LiteralPath $BasePath -Directory -ErrorAction Stop } catch { $dirs = @() }

    foreach ($d in $dirs) {
        $nameNorm = Convert-ToHalfWidthDigits -Text $d.Name
        if ($nameNorm -match $rxYear -and $nameNorm -match $rxMonth -and $nameNorm -match '月') {
            return $d.FullName
        }
    }
    return $null
}


function Select-FolderPath {
    [CmdletBinding()]
    param([string]$InitialPath)

    # Explorer 風（古いが見た目はエクスプローラ）
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.BrowseForFolder(0, 'フォルダを選択してください', 0, $InitialPath)
        if ($folder) { return $folder.Self.Path }
    } catch {}

    # フォールバック
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $InitialPath
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
    return $null
}


# --- 年度フォルダの検索：名前に yyyy が含まれていればOK（"2026年度" / "2026年" / "(2026)" etc）
function Find-YearFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][int]$Year
    )

    if (-not (Test-Path -LiteralPath $BasePath)) { return $null }

    $y = [string]$Year
    try {
        $dirs = Get-ChildItem -LiteralPath $BasePath -Directory -ErrorAction Stop
        foreach ($d in $dirs) {
            $nameNorm = Convert-ToHalfWidthDigits -Text $d.Name
            if ($nameNorm -match $y) {
                return $d.FullName
            }
        }
    } catch {
        return $null
    }

    return $null
}

# --- 月フォルダの検索：名前に月番号(1-12)が含まれていればOK（"01月"/"1月"/"1"/"1_"等）
function Find-MonthFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$YearPath,
        [Parameter(Mandatory)][int]$Month
    )

    if (-not (Test-Path -LiteralPath $YearPath)) { return $null }

    $m = [int]$Month
    if ($m -lt 1 -or $m -gt 12) { return $null }

    # 例: 1月/01月/1
    $patterns = @(
        '(?<!\d){0}(?!\d)' -f $m,
        '(?<!\d){0}(?!\d)' -f ('{0:00}' -f $m)
    )

    try {
        $dirs = Get-ChildItem -LiteralPath $YearPath -Directory -ErrorAction Stop
        foreach ($d in $dirs) {
            $nameNorm = Convert-ToHalfWidthDigits -Text $d.Name
            foreach ($pat in $patterns) {
                if ($nameNorm -match $pat) {
                    return $d.FullName
                }
            }
        }
    } catch {
        return $null
    }

    return $null
}

# --- 年度/月パスを解決（Create:$true なら必要に応じて作成）
function Resolve-YearMonthPath {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][datetime]$Date,
        [switch]$Create
    )

    $cfg = $script:FolderMakerConfig
    if ($null -eq $cfg) { $cfg = @{} }

    $useFiscal = Convert-ToBoolOrDefault -Value $cfg['UseFiscalYear'] -Default $false
    $fyStart   = [int]($cfg['FiscalYearStartMonth'] | ForEach-Object { if ($_){ $_ } else { '1' } })
    if ($fyStart -lt 1 -or $fyStart -gt 12) { $fyStart = 1 }

    $yearSuffix     = if ($cfg['YearSuffix']) { [string]$cfg['YearSuffix'] } else { '年度' }
    $monthZeroPad   = Convert-ToBoolOrDefault -Value $cfg['MonthZeroPad'] -Default $false
    $allowCombined  = Convert-ToBoolOrDefault -Value $cfg['AllowCombinedYearMonth'] -Default $true
    $searchDepth    = [int]($cfg['SearchDepthForMMdd'] | ForEach-Object { if ($_){ $_ } else { '3' } })
    if ($searchDepth -lt 0) { $searchDepth = 0 }

    $calendarYear = $Date.Year
    $month = $Date.Month
    $year = if ($useFiscal) { Get-FiscalYear -Date $Date -StartMonth $fyStart } else { $calendarYear }

    $baseInfo = Get-BasePathInfo -BasePath $BasePath
    $isBaseMonth = ($baseInfo.Kind -eq 'Month')
    $isBaseYear = ($baseInfo.Kind -eq 'Year')

    if ($isBaseMonth) {
        return [pscustomobject]@{
            Year                 = $year
            Month                = $month
            YearFolderName        = $null
            MonthFolderName       = (Split-Path -Leaf $BasePath)
            YearPath             = $null
            MonthPath            = $BasePath
            YearMatchedExisting  = $false
            MonthMatchedExisting = $true
            UsesExistingMMddParent = $true
            UsesCombined         = $false
        }
    }

    # A案：既存 "MMdd-n" を優先（親フォルダを採用） (Root 配下のみ探索)
    if (-not $isBaseYear) {
        $dateStr = $Date.ToString('MMdd')
        $mmddParent = Find-ExistingMonthContainerByMMdd -BasePath $BasePath -DateString $dateStr -MaxDepth $searchDepth -Year $calendarYear -Month $month
        if ($mmddParent) {
            return [pscustomobject]@{
                Year                 = $year
                Month                = $month
                YearFolderName        = $null
                MonthFolderName       = (Split-Path -Leaf $mmddParent)
                YearPath             = $null
                MonthPath            = $mmddParent
                YearMatchedExisting  = $false
                MonthMatchedExisting = $true
                UsesExistingMMddParent = $true
                UsesCombined         = $false
            }
        }
    }

    # 標準階層（Year\Month）を探す
    $yearExistingPath = if ($isBaseYear) { $BasePath } else { Find-YearFolder -BasePath $BasePath -Year $year }

    # FiscalYear を使っていて見つからない場合、念のためカレンダー年も試す
    if (-not $yearExistingPath -and $useFiscal -and $year -ne $calendarYear) {
        $alt = Find-YearFolder -BasePath $BasePath -Year $calendarYear
        if ($alt) { $yearExistingPath = $alt; $year = $calendarYear }
    }

    $yearFolderName = ('{0}{1}' -f $year, $yearSuffix)
    $yearPath = if ($yearExistingPath) { $yearExistingPath } else { (Join-Path $BasePath $yearFolderName) }
    $yearMatched = [bool]$yearExistingPath

    $monthExistingPath = if (Test-Path -LiteralPath $yearPath) { Find-MonthFolder -YearPath $yearPath -Month $month } else { $null }
    $monthFolderName = if ($monthZeroPad) { ('{0:00}月' -f $month) } else { ('{0}月' -f $month) }
    $standardMonthPath = if ($monthExistingPath) { $monthExistingPath } else { (Join-Path $yearPath $monthFolderName) }

    $standardMonthExists = (Test-Path -LiteralPath $standardMonthPath)

    # B案：結合フォルダ "YYYY年M月 / YYYY年度M月"（標準が無い場合のみ採用）
    $combinedPath = $null
    if ($allowCombined -and -not $standardMonthExists) {
        $combinedRoots = @($yearPath)
        if (-not $isBaseYear) { $combinedRoots = @($BasePath, $yearPath) }
        foreach ($root in $combinedRoots) {
            if (-not $root) { continue }
            $combinedPath = Find-CombinedYearMonthFolder -BasePath $root -Year $calendarYear -Month $month
            if (-not $combinedPath -and $useFiscal -and $year -ne $calendarYear) {
                $combinedPath = Find-CombinedYearMonthFolder -BasePath $root -Year $year -Month $month
            }
            if ($combinedPath) { break }
        }
    }

    $useCombined = [bool]$combinedPath
    $monthPath = if ($useCombined) { $combinedPath } else { $standardMonthPath }

    # Create 指定時のみ作成（結合フォルダ採用時は Year/Month の自動作成はしない）
    if ($Create -and -not $useCombined) {
        if (-not (Test-Path -LiteralPath $yearPath)) {
            if ($PSCmdlet.ShouldProcess($yearPath, 'Create 年度 folder')) {
                New-Item -ItemType Directory -Path $yearPath -Force -ErrorAction Stop | Out-Null
            }
        }
        if (-not (Test-Path -LiteralPath $monthPath)) {
            if ($PSCmdlet.ShouldProcess($monthPath, 'Create 月 folder')) {
                New-Item -ItemType Directory -Path $monthPath -Force -ErrorAction Stop | Out-Null
            }
        }
    }

    if (-not ($monthPath.StartsWith($BasePath, [StringComparison]::OrdinalIgnoreCase))) {
        throw ('Resolved MonthPath is outside BasePath. BasePath={0} MonthPath={1}' -f $BasePath, $monthPath)
    }

    return [pscustomobject]@{
        Year                 = $year
        Month                = $month
        YearFolderName        = $yearFolderName
        MonthFolderName       = if ($useCombined) { (Split-Path -Leaf $monthPath) } else { $monthFolderName }
        YearPath             = if ($useCombined) { $null } else { $yearPath }
        MonthPath            = $monthPath
        YearMatchedExisting  = $yearMatched
        MonthMatchedExisting = [bool]($monthExistingPath -or $useCombined)
        UsesExistingMMddParent = $false
        UsesCombined         = $useCombined
    }
}


function Get-ExistingIndicesForDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$DateString # MMdd
    )

    $pattern = ('^{0}-\d+$' -f [Regex]::Escape($DateString))

    # 何があっても HashSet を返す（null対策）
    $indices = New-Object 'System.Collections.Generic.HashSet[int]'

    if (-not (Test-Path -LiteralPath $BasePath)) {
        return $indices
    }

    try {
        $names = Get-ChildItem -LiteralPath $BasePath -Directory -Name -ErrorAction Stop
        foreach ($n in $names) {
            if ($n -match $pattern) {
                $parts = $n -split '-'
                if ($parts.Count -ge 2) {
                    $i = $parts[1] -as [int]
                    if ($null -ne $i) { [void]$indices.Add($i) }
                }
            }
        }
    }
    catch {
        # ネットワークドライブ等で一時エラーでも落とさず「既存なし」として扱う
        return $indices
    }

    return $indices
}

function Get-ExistingIndicesByMonth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$MonthPath,
        [ref]$ErrorMessage
    )

    $result = @{}
    if (-not (Test-Path -LiteralPath $MonthPath)) {
        return $result
    }

    try {
        $names = Get-ChildItem -LiteralPath $MonthPath -Directory -Name -ErrorAction Stop
        foreach ($n in $names) {
            $nameNorm = Convert-ToHalfWidthDigits -Text $n
            if ($nameNorm -match '^(?<date>\d{4})-(?<index>\d+)$') {
                $dateKey = $Matches['date']
                $index = $Matches['index'] -as [int]
                if ($null -eq $index) { continue }
                if (-not $result.ContainsKey($dateKey)) {
                    $result[$dateKey] = New-Object 'System.Collections.Generic.HashSet[int]'
                }
                [void]$result[$dateKey].Add($index)
            }
        }
    }
    catch {
        if ($ErrorMessage) { $ErrorMessage.Value = $_.Exception.Message }
        # ネットワークドライブ等の一時エラーでも落とさず空結果で返す
        return @{}
    }

    return $result
}

function Get-DateRange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][ValidateSet('Range','Days')][string]$Mode,
        [datetime]$EndDate,
        [int]$DaysToMake
    )

    switch ($Mode) {
        'Range' {
            if (-not $EndDate) { throw 'Mode=Range の場合 EndDate が必要です。' }
            if ($EndDate.Date -lt $StartDate.Date) { throw '終了日は開始日以降にしてください。' }
            return @($StartDate.Date, $EndDate.Date)
        }
        'Days' {
            if ($DaysToMake -lt 1) { throw '日数は1以上にしてください。' }
            $end = $StartDate.Date.AddDays($DaysToMake - 1)
            return @($StartDate.Date, $end)
        }
    }
}

function Get-PlannedFolderList {
    <#
      作成計画（プレビュー）だけ返す（作らない）
      追加：年度/月フォルダも計画に含める

      戻り値：
        Kind(Year/Month/Day)
        Date / Index / FolderName / FullPath / Exists / Action(Create/Skip)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [Parameter(Mandatory)][ValidateRange(1,999)][int]$FoldersPerDay,
        [Parameter(Mandatory)][ValidateRange(1,999)][int]$FirstDayStartIndex
    )

    $results = New-Object System.Collections.Generic.List[object]
    $totalDays = ($EndDate.Date - $StartDate.Date).Days + 1

    # 重複防止用
    $seenYear  = New-Object 'System.Collections.Generic.HashSet[string]'
    $seenMonth = New-Object 'System.Collections.Generic.HashSet[string]'
    $monthIndexCache = @{}
    $ymCache = @{}

    for ($dayOffset = 0; $dayOffset -lt $totalDays; $dayOffset++) {
        $date = $StartDate.Date.AddDays($dayOffset)
        $ymKey = $date.ToString('yyyyMM')
        if ($ymCache.ContainsKey($ymKey)) {
            $ym = $ymCache[$ymKey]
        } else {
            $ym = Resolve-YearMonthPath -BasePath $BasePath -Date $date
            $ymCache[$ymKey] = $ym
        }

        # 年度フォルダ計画
        if ($ym.YearPath -and -not $seenYear.Contains($ym.YearPath)) {
            [void]$seenYear.Add($ym.YearPath)

            $yearExistsPhys = (Test-Path -LiteralPath $ym.YearPath)
            $yearAction = if ($yearExistsPhys -or $ym.YearMatchedExisting) { 'Skip' } else { 'Create' }

            $results.Add([pscustomobject]@{
                Kind       = 'Year'
                Date       = $date
                Index      = $null
                FolderName = (Split-Path -Leaf $ym.YearPath)
                FullPath   = $ym.YearPath
                Exists     = $yearExistsPhys -or $ym.YearMatchedExisting
                Action     = $yearAction
            })
        }

        # 月フォルダ計画
        if (-not $seenMonth.Contains($ym.MonthPath)) {
            [void]$seenMonth.Add($ym.MonthPath)

            $monthExistsPhys = (Test-Path -LiteralPath $ym.MonthPath)
            $monthAction = if ($monthExistsPhys -or $ym.MonthMatchedExisting) { 'Skip' } else { 'Create' }

            $results.Add([pscustomobject]@{
                Kind       = 'Month'
                Date       = $date
                Index      = $null
                FolderName = (Split-Path -Leaf $ym.MonthPath)
                FullPath   = $ym.MonthPath
                Exists     = $monthExistsPhys -or $ym.MonthMatchedExisting
                Action     = $monthAction
            })
        }

        # ★変更：フォルダ名の日付部分を MMdd に
        $dateStr = $date.ToString('MMdd')

        if (-not $monthIndexCache.ContainsKey($ym.MonthPath)) {
            $monthError = $null
            $monthIndexCache[$ym.MonthPath] = Get-ExistingIndicesByMonth -MonthPath $ym.MonthPath -ErrorMessage ([ref]$monthError)
            if ($monthError) {
                throw ('月フォルダの列挙に失敗しました: {0} ({1})' -f $ym.MonthPath, $monthError)
            }
        }
        $monthIndexMap = $monthIndexCache[$ym.MonthPath]
        $existingIndices = if ($monthIndexMap.ContainsKey($dateStr)) { $monthIndexMap[$dateStr] } else { $null }
        if ($null -eq $existingIndices) {
            $existingIndices = New-Object 'System.Collections.Generic.HashSet[int]'
        }

        $startIndex = if ($dayOffset -eq 0) { $FirstDayStartIndex } else { 1 }

        for ($i = $startIndex; $i -le $FoldersPerDay; $i++) {
            $folderName = ('{0}-{1}' -f $dateStr, $i)
            $fullPath = Join-Path $ym.MonthPath $folderName

            $exists = $false
            try {
                $exists = $existingIndices.Contains($i) -or (Test-Path -LiteralPath $fullPath)
            } catch {
                $exists = $false
            }

            $results.Add([pscustomobject]@{
                Kind       = 'Day'
                Date       = $date
                Index      = $i
                FolderName = $folderName
                FullPath   = $fullPath
                Exists     = $exists
                Action     = if ($exists) { 'Skip' } else { 'Create' }
            })
        }
    }

    return $results
}

function New-DateIndexedFolders {
    <#
      実処理：フォルダ作成
      - 既存はスキップ
      - 追加：年度/月フォルダを必要に応じて作成
      - 結果一覧を返す
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [Parameter(Mandatory)][ValidateRange(1,999)][int]$FoldersPerDay,
        [Parameter(Mandatory)][ValidateRange(1,999)][int]$FirstDayStartIndex
    )

    if (-not (Test-Path -LiteralPath $BasePath)) {
        throw ('作成先パスが存在しません: {0}' -f $BasePath)
    }

    $plan = Get-PlannedFolderList -BasePath $BasePath -StartDate $StartDate -EndDate $EndDate `
        -FoldersPerDay $FoldersPerDay -FirstDayStartIndex $FirstDayStartIndex

    $out = New-Object System.Collections.Generic.List[object]

    foreach ($item in $plan) {
        if ($item.Action -eq 'Skip') {
            $out.Add([pscustomobject]@{
                FolderName = $item.FolderName
                FullPath   = $item.FullPath
                Result     = if ($item.Kind -eq 'Year') { 'Skipped(YearExists)' } elseif ($item.Kind -eq 'Month') { 'Skipped(MonthExists)' } else { 'Skipped(Exists)' }
            })
            continue
        }

        if ($PSCmdlet.ShouldProcess($item.FullPath, 'Create directory')) {
            New-Item -ItemType Directory -Path $item.FullPath -Force -ErrorAction Stop | Out-Null
            $out.Add([pscustomobject]@{
                FolderName = $item.FolderName
                FullPath   = $item.FullPath
                Result     = if ($item.Kind -eq 'Year') { 'Created(Year)' } elseif ($item.Kind -eq 'Month') { 'Created(Month)' } else { 'Created' }
            })
        }
    }

    return $out
}


Export-ModuleMember -Function Initialize-FolderMakerConfig,Get-FolderMakerConfigPath,Load-FolderMakerConfig,Save-FolderMakerConfig,Get-BasePathInfo,Resolve-YearMonthPath,Get-DateRange,Get-PlannedFolderList,New-DateIndexedFolders
