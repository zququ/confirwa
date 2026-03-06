param(
    [string] $StateFile = "",
    [string] $ImageDir = "D:\EdgeDownload\chiikawa",
    [double] $IdleSeconds = 1.5,
    [int] $PollMs = 400
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($StateFile)) {
    $StateFile = Join-Path $env:USERPROFILE ".cache\confirwa.state"
}

function Ensure-ParentDir {
    param([string] $Path)
    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path $parent)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
}

function Read-State {
    if (-not (Test-Path $StateFile)) {
        return "idle"
    }
    try {
        $raw = Get-Content -Path $StateFile -TotalCount 1 -ErrorAction Stop
    } catch {
        return "idle"
    }
    $state = ""
    if ($null -ne $raw) {
        $state = $raw.ToString().Trim().ToLowerInvariant()
    }
    switch ($state) {
        "running" { return "running" }
        "approval" { return "approval" }
        "reconnecting" { return "reconnecting" }
        "silent" { return "silent" }
        "idle" { return "idle" }
        default { return "idle" }
    }
}

function Write-State {
    param([string] $State)
    switch ($State) {
        "running" {}
        "approval" {}
        "reconnecting" {}
        "silent" {}
        "idle" {}
        default { return }
    }
    Ensure-ParentDir $StateFile
    try {
        Set-Content -Path $StateFile -Value $State -Encoding ascii
    } catch {
    }
}

function Get-ImagePath {
    param([string] $State)
    switch ($State) {
        "running" { return (Join-Path $ImageDir "1giphy.gif") }
        "approval" { return (Join-Path $ImageDir "2giphy.gif") }
        "reconnecting" { return (Join-Path $ImageDir "3giphy.gif") }
        "silent" { return (Join-Path $ImageDir "5giphy.gif") }
        default { return (Join-Path $ImageDir "4giphy.gif") }
    }
}

function Get-BaseNameFromPath {
    param([string] $Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return "codex"
    }
    $trim = $Path.Trim().Trim('"').TrimEnd('\', '/')
    if ([string]::IsNullOrWhiteSpace($trim)) {
        return "codex"
    }
    $leaf = Split-Path -Leaf $trim
    if ([string]::IsNullOrWhiteSpace($leaf)) {
        $leaf = $trim
    }
    if ([string]::IsNullOrWhiteSpace($leaf)) {
        return "codex"
    }
    return $leaf
}

$script:currentState = Read-State
$script:sessionMetaCache = @{}
$script:sessionOffsets = @{}
$script:sessionStatus = @{}
$script:recentSessionInfos = @()
$script:lastSessionRefreshUtc = [DateTime]::MinValue
$script:errorLog = Join-Path $env:USERPROFILE ".cache\confirwa-error.log"

$script:lastCodexProbeUtc = [DateTime]::MinValue
$script:lastCodexCount = 0
$script:codexStateDb = Join-Path $env:USERPROFILE ".codex\state_5.sqlite"
$script:lastActivePathProbeUtc = [DateTime]::MinValue
$script:activeRolloutPathSet = @{}
$script:lastReconnectProbeUtc = [DateTime]::MinValue
$script:lastReconnectLogId = -1
$script:lastTerminalHwnd = [IntPtr]::Zero
$script:driverForm = $null
$script:tabMapFile = Join-Path $env:USERPROFILE ".cache\confirwa-tabmap.json"
$script:titleTabMap = @{}
$script:orderFile = Join-Path $env:USERPROFILE ".cache\confirwa-order.json"
$script:slotOrder = @()
$script:anchorFile = Join-Path $env:USERPROFILE ".cache\confirwa-anchor.json"
$script:anchorOffsetX = 0
$script:anchorOffsetY = 0
$script:groupDragState = $null
$script:dragState = @{}
$script:clickSuppressUntil = @{}

function Set-CurrentState {
    param([string] $State)
    switch ($State) {
        "running" {}
        "approval" {}
        "reconnecting" {}
        "silent" {}
        "idle" {}
        default { return }
    }
    if ($script:currentState -ne $State) {
        $script:currentState = $State
        Write-State $State
    }
}

function Normalize-TabTitleKey {
    param([string] $Title)
    if ([string]::IsNullOrWhiteSpace($Title)) {
        return "codex"
    }
    $t = [string]$Title
    $t = $t.Trim()
    $t = $t -replace '\s+\(\d+\)$', ''
    if ([string]::IsNullOrWhiteSpace($t)) {
        $t = "codex"
    }
    return $t.ToLowerInvariant()
}

function Save-TabMap {
    try {
        Ensure-ParentDir $script:tabMapFile
        $obj = [ordered]@{}
        foreach ($k in @($script:titleTabMap.Keys | Sort-Object)) {
            $v = [int]$script:titleTabMap[$k]
            if ($v -ge 1 -and $v -le 9) {
                $obj[$k] = $v
            }
        }
        ($obj | ConvertTo-Json -Depth 3) | Set-Content -Path $script:tabMapFile -Encoding UTF8
    } catch {
    }
}

function Load-TabMap {
    $script:titleTabMap = @{}
    if (-not (Test-Path $script:tabMapFile)) {
        return
    }
    try {
        $raw = Get-Content -Path $script:tabMapFile -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return
        }
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($obj -is [System.Collections.IDictionary]) {
            foreach ($k in $obj.Keys) {
                $nk = Normalize-TabTitleKey -Title ([string]$k)
                $v = [int]$obj[$k]
                if ($v -ge 1 -and $v -le 9) {
                    $script:titleTabMap[$nk] = $v
                }
            }
            return
        }
        foreach ($prop in @($obj.PSObject.Properties)) {
            $nk = Normalize-TabTitleKey -Title ([string]$prop.Name)
            $v = [int]$prop.Value
            if ($v -ge 1 -and $v -le 9) {
                $script:titleTabMap[$nk] = $v
            }
        }
    } catch {
        $script:titleTabMap = @{}
    }
}

function Normalize-ModelKey {
    param([string] $Key)
    if ([string]::IsNullOrWhiteSpace($Key)) {
        return ""
    }
    return $Key.Trim().ToLowerInvariant()
}

function Normalize-SessionPath {
    param([string] $Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }
    $p = [string]$Path
    $p = $p.Trim().Trim('"')
    $p = $p.Replace('/', '\')
    if ($p.StartsWith("\\?\\")) {
        $p = $p.Substring(4)
    }
    return $p.ToLowerInvariant()
}

function Get-CodexProcessIds {
    $ids = @()
    try {
        $ids = @(
            Get-Process -Name codex -ErrorAction SilentlyContinue |
                Sort-Object Id |
                ForEach-Object { [int]$_.Id }
        )
    } catch {
        $ids = @()
    }
    return @($ids)
}

function Refresh-ActiveRolloutPathSet {
    $now = [DateTime]::UtcNow
    if (($now - $script:lastActivePathProbeUtc).TotalMilliseconds -lt 1200) {
        return
    }
    $script:lastActivePathProbeUtc = $now
    $script:activeRolloutPathSet = @{}

    if (-not (Test-Path $script:codexStateDb)) {
        return
    }

    $sqlite = Get-Command sqlite3 -ErrorAction SilentlyContinue
    if (-not $sqlite) {
        return
    }

    $pids = Get-CodexProcessIds
    if ($pids.Count -eq 0) {
        return
    }

    $want = @{}
    foreach ($procId in $pids) {
        if ($procId -gt 0) {
            $want[[string]$procId] = $true
        }
    }
    if ($want.Count -eq 0) {
        return
    }

    $pidThread = @{}
    $logRows = @()
    $minTs = [int64]([DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - 172800)
    try {
        $queryLogs = "SELECT process_uuid, thread_id, ts FROM logs WHERE thread_id IS NOT NULL ORDER BY id DESC LIMIT 2500;"
        $logRows = @(& sqlite3 $script:codexStateDb $queryLogs 2>$null)
    } catch {
        $logRows = @()
    }

    foreach ($row in $logRows) {
        if ($pidThread.Count -ge $want.Count) {
            break
        }
        if ([string]::IsNullOrWhiteSpace($row)) {
            continue
        }
        $parts = [string]$row -split '\|', 3
        if ($parts.Count -lt 2) {
            continue
        }
        $procUuid = [string]$parts[0]
        $threadId = [string]$parts[1]
        if ($parts.Count -ge 3) {
            $ts = 0
            try { $ts = [int64]$parts[2] } catch { $ts = 0 }
            if ($ts -gt 0 -and $ts -lt $minTs) {
                continue
            }
        }
        if ([string]::IsNullOrWhiteSpace($procUuid) -or [string]::IsNullOrWhiteSpace($threadId)) {
            continue
        }
        $m = [regex]::Match($procUuid, '^pid:(\d+):')
        if (-not $m.Success) {
            continue
        }
        $pidKey = [string]$m.Groups[1].Value
        if (-not $want.ContainsKey($pidKey)) {
            continue
        }
        if (-not $pidThread.ContainsKey($pidKey)) {
            $pidThread[$pidKey] = $threadId
        }
    }

    if ($pidThread.Count -eq 0) {
        return
    }

    $tidSeen = @{}
    $quoted = @()
    foreach ($pidKey in @($pidThread.Keys)) {
        $tid = [string]$pidThread[$pidKey]
        if ([string]::IsNullOrWhiteSpace($tid)) {
            continue
        }
        if ($tidSeen.ContainsKey($tid)) {
            continue
        }
        $tidSeen[$tid] = $true
        $quoted += ("'" + ($tid -replace "'", "''") + "'")
    }
    if ($quoted.Count -eq 0) {
        return
    }

    $threadRows = @()
    try {
        $inList = [string]::Join(",", $quoted)
        $queryThreads = "SELECT id, rollout_path FROM threads WHERE id IN (" + $inList + ");"
        $threadRows = @(& sqlite3 $script:codexStateDb $queryThreads 2>$null)
    } catch {
        $threadRows = @()
    }

    foreach ($row in $threadRows) {
        if ([string]::IsNullOrWhiteSpace($row)) {
            continue
        }
        $parts = [string]$row -split '\|', 2
        if ($parts.Count -lt 2) {
            continue
        }
        $path = Normalize-SessionPath -Path ([string]$parts[1])
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            $script:activeRolloutPathSet[$path] = $true
        }
    }
}

function Refresh-ReconnectSignalsFromCodexLogs {
    $now = [DateTime]::UtcNow
    if (($now - $script:lastReconnectProbeUtc).TotalMilliseconds -lt 1200) {
        return
    }
    $script:lastReconnectProbeUtc = $now

    if (-not (Test-Path $script:codexStateDb)) {
        return
    }
    $sqlite = Get-Command sqlite3 -ErrorAction SilentlyContinue
    if (-not $sqlite) {
        return
    }

    if ([int64]$script:lastReconnectLogId -lt 0) {
        try {
            $seedRows = @(& sqlite3 $script:codexStateDb "SELECT IFNULL(MAX(id)-500,0) FROM logs;" 2>$null)
            if ($seedRows.Count -gt 0) {
                $seed = 0
                try { $seed = [int64]([string]$seedRows[0]).Trim() } catch { $seed = 0 }
                if ($seed -lt 0) { $seed = 0 }
                $script:lastReconnectLogId = $seed
            } else {
                $script:lastReconnectLogId = 0
            }
        } catch {
            $script:lastReconnectLogId = 0
        }
    }

    $query = "SELECT json_object('id',l.id,'ts',l.ts,'thread_id',ifnull(l.thread_id,''),'message',ifnull(l.message,''),'rollout_path',ifnull(t.rollout_path,'')) FROM logs l LEFT JOIN threads t ON t.id=l.thread_id WHERE l.id > {0} ORDER BY l.id ASC LIMIT 1200;" -f ([int64]$script:lastReconnectLogId)
    $rows = @()
    try {
        $rows = @(& sqlite3 $script:codexStateDb $query 2>$null)
    } catch {
        $rows = @()
    }
    if ($rows.Count -eq 0) {
        return
    }

    $maxId = [int64]$script:lastReconnectLogId
    foreach ($row in $rows) {
        if ([string]::IsNullOrWhiteSpace($row)) {
            continue
        }
        $obj = $null
        try {
            $obj = $row | ConvertFrom-Json -ErrorAction Stop
        } catch {
            continue
        }
        if (-not $obj) {
            continue
        }

        $id = 0
        try { $id = [int64]$obj.id } catch { $id = 0 }
        if ($id -gt $maxId) {
            $maxId = $id
        }

        $msg = ""
        try { $msg = [string]$obj.message } catch { $msg = "" }
        $isReconnectRecovered = Is-ReconnectRecoveredTextLikeBark -Line $msg
        $isReconnect = $false
        if (-not $isReconnectRecovered) {
            $isReconnect = Is-ReconnectTextLikeBark -Line $msg
        }
        if (-not $isReconnect -and -not $isReconnectRecovered) {
            continue
        }

        $rawPath = ""
        try { $rawPath = [string]$obj.rollout_path } catch { $rawPath = "" }
        if ([string]::IsNullOrWhiteSpace($rawPath)) {
            continue
        }
        $path = Normalize-SessionPath -Path $rawPath
        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        $atUtc = [DateTime]::UtcNow
        $tsSec = 0
        try { $tsSec = [int64]$obj.ts } catch { $tsSec = 0 }
        if ($tsSec -gt 0) {
            try {
                $atUtc = [DateTimeOffset]::FromUnixTimeSeconds($tsSec).UtcDateTime
            } catch {
                $atUtc = [DateTime]::UtcNow
            }
        }
        if (([DateTime]::UtcNow - $atUtc).TotalSeconds -gt 300.0) {
            continue
        }

        if ($isReconnectRecovered) {
            Clear-SessionReconnectPending -Path $path
            Mark-SessionEvent -Path $path -AtUtc $atUtc
            continue
        }

        Mark-SessionEvent -Path $path -Reconnecting -AtUtc $atUtc
    }

    $script:lastReconnectLogId = [int64]$maxId
}

function Save-SlotOrder {
    try {
        Ensure-ParentDir $script:orderFile
        $arr = @()
        foreach ($k in @($script:slotOrder)) {
            $nk = Normalize-ModelKey -Key ([string]$k)
            if (-not [string]::IsNullOrWhiteSpace($nk) -and -not $nk.StartsWith("__empty_")) {
                $arr += $nk
            }
        }
        ($arr | ConvertTo-Json -Depth 3) | Set-Content -Path $script:orderFile -Encoding UTF8
    } catch {
    }
}

function Load-SlotOrder {
    $script:slotOrder = @()
    if (-not (Test-Path $script:orderFile)) {
        return
    }
    try {
        $raw = Get-Content -Path $script:orderFile -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return
        }
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $arr = @()
        if ($obj -is [System.Array]) {
            foreach ($v in $obj) {
                $nk = Normalize-ModelKey -Key ([string]$v)
                if (-not [string]::IsNullOrWhiteSpace($nk) -and -not $nk.StartsWith("__empty_")) {
                    $arr += $nk
                }
            }
        } elseif ($obj) {
            $nk = Normalize-ModelKey -Key ([string]$obj)
            if (-not [string]::IsNullOrWhiteSpace($nk) -and -not $nk.StartsWith("__empty_")) {
                $arr += $nk
            }
        }
        $script:slotOrder = @($arr)
    } catch {
        $script:slotOrder = @()
    }
}

function Save-AnchorOffset {
    try {
        Ensure-ParentDir $script:anchorFile
        $obj = [ordered]@{
            X = [int]$script:anchorOffsetX
            Y = [int]$script:anchorOffsetY
        }
        ($obj | ConvertTo-Json -Depth 3) | Set-Content -Path $script:anchorFile -Encoding UTF8
    } catch {
    }
}

function Load-AnchorOffset {
    $script:anchorOffsetX = 0
    $script:anchorOffsetY = 0
    if (-not (Test-Path $script:anchorFile)) {
        return
    }
    try {
        $raw = Get-Content -Path $script:anchorFile -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return
        }
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        try { $script:anchorOffsetX = [int]$obj.X } catch { $script:anchorOffsetX = 0 }
        try { $script:anchorOffsetY = [int]$obj.Y } catch { $script:anchorOffsetY = 0 }
        $script:anchorOffsetX = [Math]::Max(-4000, [Math]::Min(4000, [int]$script:anchorOffsetX))
        $script:anchorOffsetY = [Math]::Max(-4000, [Math]::Min(4000, [int]$script:anchorOffsetY))
    } catch {
        $script:anchorOffsetX = 0
        $script:anchorOffsetY = 0
    }
}

function Apply-SlotOrder {
    param(
        [object[]] $Models,
        [int] $Count
    )

    if ($null -eq $Models) {
        return @()
    }
    $input = @($Models)
    if ($input.Count -eq 0) {
        return @()
    }

    $picked = @()
    $used = @{}

    foreach ($k in @($script:slotOrder)) {
        $nk = Normalize-ModelKey -Key ([string]$k)
        if ([string]::IsNullOrWhiteSpace($nk)) {
            continue
        }
        if ($used.ContainsKey($nk)) {
            continue
        }
        foreach ($m in $input) {
            $mk = Normalize-ModelKey -Key ([string]$m.Key)
            if ($mk -eq $nk) {
                $picked += $m
                $used[$mk] = $true
                break
            }
        }
    }

    foreach ($m in $input) {
        $mk = Normalize-ModelKey -Key ([string]$m.Key)
        if ([string]::IsNullOrWhiteSpace($mk)) {
            continue
        }
        if ($used.ContainsKey($mk)) {
            continue
        }
        $picked += $m
        $used[$mk] = $true
    }

    if ($Count -gt 0 -and $picked.Count -gt $Count) {
        $picked = @($picked | Select-Object -First $Count)
    }

    $newOrder = @()
    foreach ($m in $picked) {
        $mk = Normalize-ModelKey -Key ([string]$m.Key)
        if ([string]::IsNullOrWhiteSpace($mk) -or $mk.StartsWith("__empty_")) {
            continue
        }
        if (-not ($newOrder -contains $mk)) {
            $newOrder += $mk
        }
    }

    $changed = $true
    if ($script:slotOrder.Count -eq $newOrder.Count) {
        $changed = $false
        for ($i = 0; $i -lt $newOrder.Count; $i++) {
            if ([string]$script:slotOrder[$i] -ne [string]$newOrder[$i]) {
                $changed = $true
                break
            }
        }
    }
    if ($changed) {
        $script:slotOrder = @($newOrder)
        Save-SlotOrder
    }

    return @($picked)
}

function Set-ClickSuppress {
    param(
        [System.Windows.Forms.Form] $Form,
        [int] $Milliseconds = 280
    )
    if (-not $Form) {
        return
    }
    $key = [string]$Form.Handle.ToInt64()
    $script:clickSuppressUntil[$key] = [DateTime]::UtcNow.AddMilliseconds([Math]::Max(50, $Milliseconds)).Ticks
}

function Is-ClickSuppressed {
    param([System.Windows.Forms.Form] $Form)
    if (-not $Form) {
        return $false
    }
    $key = [string]$Form.Handle.ToInt64()
    if (-not $script:clickSuppressUntil.ContainsKey($key)) {
        return $false
    }
    $ticks = 0
    try { $ticks = [int64]$script:clickSuppressUntil[$key] } catch { $ticks = 0 }
    if ($ticks -le 0) {
        $script:clickSuppressUntil.Remove($key) | Out-Null
        return $false
    }
    if ([DateTime]::UtcNow.Ticks -le $ticks) {
        return $true
    }
    $script:clickSuppressUntil.Remove($key) | Out-Null
    return $false
}

function Resolve-TabHotkeyIndex {
    param(
        [string] $Title,
        [int] $Preferred = 0,
        [int] $MaxIndex = 9,
        [hashtable] $Occupied = $null
    )

    if ($MaxIndex -lt 1) {
        $MaxIndex = 1
    }
    if ($MaxIndex -gt 9) {
        $MaxIndex = 9
    }

    $key = Normalize-TabTitleKey -Title $Title
    $existing = 0
    if ($script:titleTabMap.ContainsKey($key)) {
        $existing = [int]$script:titleTabMap[$key]
    }
    if ($existing -ge 1 -and $existing -le $MaxIndex) {
        if ($null -eq $Occupied -or -not $Occupied.ContainsKey($existing)) {
            if ($null -ne $Occupied) {
                $Occupied[$existing] = $true
            }
            return $existing
        }
    }

    if ($Preferred -ge 1 -and $Preferred -le $MaxIndex) {
        if ($null -eq $Occupied -or -not $Occupied.ContainsKey($Preferred)) {
            $script:titleTabMap[$key] = $Preferred
            if ($null -ne $Occupied) {
                $Occupied[$Preferred] = $true
            }
            Save-TabMap
            return $Preferred
        }
    }

    for ($i = 1; $i -le $MaxIndex; $i++) {
        if ($null -ne $Occupied -and $Occupied.ContainsKey($i)) {
            continue
        }
        $script:titleTabMap[$key] = $i
        if ($null -ne $Occupied) {
            $Occupied[$i] = $true
        }
        Save-TabMap
        return $i
    }

    return [Math]::Max(1, [Math]::Min($Preferred, $MaxIndex))
}

function Write-ConfirwaError {
    param([string] $Message)
    if ([string]::IsNullOrWhiteSpace($Message)) {
        return
    }
    try {
        Ensure-ParentDir $script:errorLog
        Add-Content -Path $script:errorLog -Value ("{0} {1}" -f (Get-Date).ToString("o"), $Message) -Encoding UTF8
    } catch {
    }
}

function Get-CwdFromTurnContextLine {
    param([string] $Line)
    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $null
    }
    if ($Line -notmatch '"type"\s*:\s*"turn_context"') {
        return $null
    }
    $m = [regex]::Match($Line, '"cwd"\s*:\s*"([^"]+)"')
    if (-not $m.Success) {
        return $null
    }
    $cwd = $m.Groups[1].Value
    if ([string]::IsNullOrWhiteSpace($cwd)) {
        return $null
    }
    $cwd = $cwd.Replace('\\\\', '\')
    return $cwd
}

function Read-SessionMetaCwd {
    param([string] $Path)
    $fs = $null
    $sr = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true, 4096, $true)
        if (-not $sr.EndOfStream) {
            $line = $sr.ReadLine()
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $obj = $line | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($obj -and $obj.type -eq "session_meta" -and $obj.payload) {
                    $cwd = [string]$obj.payload.cwd
                    if (-not [string]::IsNullOrWhiteSpace($cwd)) {
                        return $cwd
                    }
                }
            }
        }
    } catch {
    } finally {
        if ($sr) { $sr.Dispose() }
        if ($fs) { $fs.Dispose() }
    }
    return $null
}

function Read-RecentTurnContextCwd {
    param(
        [string] $Path,
        [int64] $Length
    )

    if ($Length -le 0) {
        return $null
    }

    $tailBytes = [int64]524288
    $start = [Math]::Max([int64]0, $Length - $tailBytes)

    $fs = $null
    $sr = $null
    $lastCwd = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        [void]$fs.Seek($start, [System.IO.SeekOrigin]::Begin)
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true, 4096, $true)
        if ($start -gt 0 -and -not $sr.EndOfStream) {
            [void]$sr.ReadLine()
        }
        while (-not $sr.EndOfStream) {
            $line = $sr.ReadLine()
            $cwd = Get-CwdFromTurnContextLine -Line $line
            if (-not [string]::IsNullOrWhiteSpace($cwd)) {
                $lastCwd = $cwd
            }
        }
    } catch {
    } finally {
        if ($sr) { $sr.Dispose() }
        if ($fs) { $fs.Dispose() }
    }

    return $lastCwd
}

function Read-SessionTitle {
    param([System.IO.FileInfo] $Item)
    if (-not $Item) {
        return "codex"
    }

    $path = [string]$Item.FullName
    if ([string]::IsNullOrWhiteSpace($path)) {
        return "codex"
    }

    $ticks = [int64]$Item.LastWriteTimeUtc.Ticks
    $length = [int64]$Item.Length

    if ($script:sessionMetaCache.ContainsKey($path)) {
        $cached = $script:sessionMetaCache[$path]
        if ($cached -and $cached.LastWriteTicks -eq $ticks -and $cached.Length -eq $length) {
            return [string]$cached.Title
        }
    }

    $cwd = $null
    $ageSec = ([DateTime]::UtcNow - $Item.LastWriteTimeUtc).TotalSeconds
    if ($ageSec -le 86400) {
        $cwd = Read-RecentTurnContextCwd -Path $path -Length $length
    }
    if ([string]::IsNullOrWhiteSpace($cwd)) {
        $cwd = Read-SessionMetaCwd -Path $path
    }

    $title = "codex"
    if (-not [string]::IsNullOrWhiteSpace($cwd)) {
        $title = Get-BaseNameFromPath $cwd
    }
    if ([string]::IsNullOrWhiteSpace($title)) {
        $title = "codex"
    }

    $script:sessionMetaCache[$path] = [PSCustomObject]@{
        Title = $title
        LastWriteTicks = $ticks
        Length = $length
    }
    return $title
}

function Get-RecentSessionDirs {
    $root = Join-Path $env:USERPROFILE ".codex\sessions"
    if (-not (Test-Path $root)) {
        return @()
    }

    $dirs = @()
    for ($i = 0; $i -le 2; $i++) {
        $dt = (Get-Date).AddMonths(-$i)
        $dir = Join-Path (Join-Path $root $dt.ToString("yyyy")) $dt.ToString("MM")
        if (Test-Path $dir) {
            $dirs += $dir
        }
    }
    if ($dirs.Count -eq 0) {
        $dirs += $root
    }
    return @($dirs)
}

function Ensure-SessionStatus {
    param([string] $Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    if (-not $script:sessionStatus.ContainsKey($Path)) {
        $script:sessionStatus[$Path] = [PSCustomObject]@{
            LastEventUtc = [DateTime]::MinValue
            LastApprovalUtc = [DateTime]::MinValue
            LastReconnectUtc = [DateTime]::MinValue
            LastWorkUtc = [DateTime]::MinValue
            LastTerminalUtc = [DateTime]::MinValue
            LastSpeechUtc = [DateTime]::MinValue
            LastSpeechText = ""
            ApprovalPending = $false
            ReconnectPending = $false
            ActiveTurns = @{}
        }
    }
}

function Mark-SessionEvent {
    param(
        [string] $Path,
        [switch] $Approval,
        [switch] $Reconnecting,
        [DateTime] $AtUtc = [DateTime]::MinValue
    )
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    Ensure-SessionStatus -Path $Path

    $when = $AtUtc
    if ($when -eq [DateTime]::MinValue) {
        $when = [DateTime]::UtcNow
    } else {
        $when = $when.ToUniversalTime()
    }

    if ($when -gt $script:sessionStatus[$Path].LastEventUtc) {
        $script:sessionStatus[$Path].LastEventUtc = $when
    }
    if ($Approval) {
        if ($when -gt $script:sessionStatus[$Path].LastApprovalUtc) {
            $script:sessionStatus[$Path].LastApprovalUtc = $when
        }
        $script:sessionStatus[$Path].ApprovalPending = $true
        $script:sessionStatus[$Path].ReconnectPending = $false
    }
    if ($Reconnecting) {
        if ($when -gt $script:sessionStatus[$Path].LastReconnectUtc) {
            $script:sessionStatus[$Path].LastReconnectUtc = $when
        }
        $script:sessionStatus[$Path].ReconnectPending = $true
    }
}

function Mark-SessionWork {
    param(
        [string] $Path,
        [DateTime] $AtUtc = [DateTime]::MinValue
    )
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    Ensure-SessionStatus -Path $Path

    $when = $AtUtc
    if ($when -eq [DateTime]::MinValue) {
        $when = [DateTime]::UtcNow
    } else {
        $when = $when.ToUniversalTime()
    }

    if ($when -gt $script:sessionStatus[$Path].LastWorkUtc) {
        $script:sessionStatus[$Path].LastWorkUtc = $when
    }
    if ($when -gt $script:sessionStatus[$Path].LastEventUtc) {
        $script:sessionStatus[$Path].LastEventUtc = $when
    }
}

function Mark-SessionTerminal {
    param(
        [string] $Path,
        [DateTime] $AtUtc = [DateTime]::MinValue
    )
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    Ensure-SessionStatus -Path $Path

    $when = $AtUtc
    if ($when -eq [DateTime]::MinValue) {
        $when = [DateTime]::UtcNow
    } else {
        $when = $when.ToUniversalTime()
    }

    if ($when -gt $script:sessionStatus[$Path].LastTerminalUtc) {
        $script:sessionStatus[$Path].LastTerminalUtc = $when
    }
    if ($when -gt $script:sessionStatus[$Path].LastEventUtc) {
        $script:sessionStatus[$Path].LastEventUtc = $when
    }
}

function Clear-SessionApprovalPending {
    param([string] $Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    Ensure-SessionStatus -Path $Path
    $script:sessionStatus[$Path].ApprovalPending = $false
}

function Clear-SessionReconnectPending {
    param([string] $Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    Ensure-SessionStatus -Path $Path
    $script:sessionStatus[$Path].ReconnectPending = $false
}

function Clear-SessionActiveTurns {
    param([string] $Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    Ensure-SessionStatus -Path $Path
    $script:sessionStatus[$Path].ActiveTurns = @{}
}

function Is-FinalAnswerSignal {
    param([string] $Line)
    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $false
    }
    $s = $Line.ToLowerInvariant()
    if ($s -match '"phase"\s*:\s*"final_answer"') { return $true }
    if ($s -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"task_complete"') { return $true }
    return $false
}

function Is-TerminalActivitySignal {
    param(
        $Obj,
        [string] $Line
    )

    if (Is-FinalAnswerSignal -Line $Line) {
        return $true
    }
    if (-not $Obj) {
        return $false
    }

    $type = ""
    try { $type = [string]$Obj.type } catch { $type = "" }
    if ($type -ne "event_msg") {
        return $false
    }

    $eventType = ""
    try { $eventType = [string]$Obj.payload.type } catch { $eventType = "" }
    return ($eventType -match "^(task_complete|turn_aborted|task_failed|task_cancelled)$")
}

function Is-WorkActivitySignal {
    param($Obj)

    if (-not $Obj) {
        return $false
    }

    $type = ""
    try { $type = [string]$Obj.type } catch { $type = "" }
    if ($type -eq "event_msg") {
        $eventType = ""
        try { $eventType = [string]$Obj.payload.type } catch { $eventType = "" }
        return ($eventType -match "^(task_started|agent_message|approval_request|requires_approval)$")
    }
    if ($type -eq "response_item") {
        $payloadType = ""
        $role = ""
        try { $payloadType = [string]$Obj.payload.type } catch { $payloadType = "" }
        try { $role = [string]$Obj.payload.role } catch { $role = "" }
        if (($payloadType -eq "message" -and $role -eq "assistant") -or
            $payloadType -eq "function_call" -or
            $payloadType -eq "function_call_output" -or
            $payloadType -eq "custom_tool_call_output") {
            return $true
        }
    }
    return $false
}

function Decode-JsonStringValue {
    param([string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }
    $decoded = ""
    try {
        $decoded = ('"' + $Value + '"') | ConvertFrom-Json -ErrorAction Stop
    } catch {
        try {
            $decoded = [System.Text.RegularExpressions.Regex]::Unescape($Value)
        } catch {
            $decoded = $Value
        }
    }
    return [string]$decoded
}

function Normalize-SpeechText {
    param(
        [string] $Text,
        [int] $MaxLength = 0
    )
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $s = [string]$Text
    $s = $s -replace '\r\n?', "`n"
    $s = $s -replace '\[(.*?)\]\([^)]+\)', '$1'
    $s = $s -replace '(?m)^\s*[-*]\s+', ''
    $s = $s -replace '`', ""
    $s = $s -replace "`t", " "
    $s = ($s -split "`n" | ForEach-Object { ($_ -replace '[ ]{2,}', ' ').Trim() }) -join "`n"
    $s = $s -replace "`n{2,}", "`n"
    $s = $s.Trim()

    if ($MaxLength -gt 0 -and $s.Length -gt $MaxLength) {
        $s = $s.Substring(0, [Math]::Max(0, $MaxLength - 3)).TrimEnd() + "..."
    }
    return $s
}

function Get-FirstSentencesText {
    param(
        [string] $Text,
        [int] $SentenceCount = 2
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }
    if ($SentenceCount -le 0) {
        return ""
    }

    $n = Normalize-SpeechText -Text $Text
    if ([string]::IsNullOrWhiteSpace($n)) {
        return ""
    }

    $firstLine = ""
    foreach ($ln in ($n -split "`n")) {
        if (-not [string]::IsNullOrWhiteSpace($ln)) {
            $firstLine = $ln.Trim()
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($firstLine)) {
        return ""
    }

    $matches = [regex]::Matches($firstLine, '[^。！？.!?]+[。！？.!?]?')
    $parts = @()
    foreach ($m in $matches) {
        $seg = [string]$m.Value
        if ([string]::IsNullOrWhiteSpace($seg)) {
            continue
        }
        $seg = $seg.Trim()
        if ($seg.Length -le 0) {
            continue
        }
        $parts += $seg
        if ($parts.Count -ge $SentenceCount) {
            break
        }
    }

    if ($parts.Count -gt 0) {
        return ([string]::Join(" ", $parts)).Trim()
    }
    return $firstLine.Trim()
}

function Extract-AssistantSpeechFromLine {
    param([string] $Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return ""
    }
    if ($Line -notmatch '"type"\s*:\s*"(event_msg|response_item)"') {
        return ""
    }

    $encoded = ""
    if ($Line -match '"type"\s*:\s*"event_msg"' -and $Line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"agent_message"') {
        $m = [regex]::Match($Line, '"message"\s*:\s*"((?:\\.|[^"\\])*)"')
        if ($m.Success) {
            $encoded = [string]$m.Groups[1].Value
        }
    }
    if ([string]::IsNullOrWhiteSpace($encoded) -and $Line -match '"type"\s*:\s*"event_msg"' -and $Line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"task_complete"') {
        $m = [regex]::Match($Line, '"last_agent_message"\s*:\s*"((?:\\.|[^"\\])*)"')
        if ($m.Success) {
            $encoded = [string]$m.Groups[1].Value
        }
    }
    if ([string]::IsNullOrWhiteSpace($encoded) -and $Line -match '"type"\s*:\s*"response_item"' -and $Line -match '"role"\s*:\s*"assistant"') {
        $m = [regex]::Match($Line, '"type"\s*:\s*"output_text"\s*,\s*"text"\s*:\s*"((?:\\.|[^"\\])*)"')
        if (-not $m.Success) {
            $m = [regex]::Match($Line, '"text"\s*:\s*"((?:\\.|[^"\\])*)"')
        }
        if ($m.Success) {
            $encoded = [string]$m.Groups[1].Value
        }
    }

    if ([string]::IsNullOrWhiteSpace($encoded)) {
        return ""
    }

    $decoded = Decode-JsonStringValue -Value $encoded
    return (Normalize-SpeechText -Text $decoded)
}

function Mark-SessionSpeech {
    param(
        [string] $Path,
        [string] $Speech,
        [DateTime] $AtUtc = [DateTime]::MinValue
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Speech)) {
        return
    }

    Ensure-SessionStatus -Path $Path
    $when = $AtUtc
    if ($when -eq [DateTime]::MinValue) {
        $when = [DateTime]::UtcNow
    } else {
        $when = $when.ToUniversalTime()
    }

    $normalized = Normalize-SpeechText -Text $Speech
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return
    }

    $status = $script:sessionStatus[$Path]
    if ($when -ge $status.LastSpeechUtc -or [string]::IsNullOrWhiteSpace([string]$status.LastSpeechText)) {
        $status.LastSpeechUtc = $when
        $status.LastSpeechText = $normalized
    }
}

function Get-SessionSpeechPreview {
    param([string] $Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return "..."
    }
    if ($script:sessionStatus.ContainsKey($Path)) {
        $t = [string]$script:sessionStatus[$Path].LastSpeechText
        if (-not [string]::IsNullOrWhiteSpace($t)) {
            $preview = Get-FirstSentencesText -Text $t -SentenceCount 1
            if (-not [string]::IsNullOrWhiteSpace($preview)) {
                return $preview
            }
        }
    }
    return "..."
}

function Get-EventTimestampUtc {
    param($Obj)
    if (-not $Obj) {
        return [DateTime]::UtcNow
    }
    $ts = ""
    try {
        $ts = [string]$Obj.timestamp
    } catch {
        $ts = ""
    }
    if ([string]::IsNullOrWhiteSpace($ts)) {
        return [DateTime]::UtcNow
    }
    try {
        return ([DateTimeOffset]::Parse($ts)).UtcDateTime
    } catch {
        return [DateTime]::UtcNow
    }
}

function Get-LineTimestampUtc {
    param([string] $Line)
    if ([string]::IsNullOrWhiteSpace($Line)) {
        return [DateTime]::UtcNow
    }
    $m = [regex]::Match($Line, '"timestamp"\s*:\s*"([^"]+)"')
    if (-not $m.Success) {
        return [DateTime]::UtcNow
    }
    try {
        return ([DateTimeOffset]::Parse($m.Groups[1].Value)).UtcDateTime
    } catch {
        return [DateTime]::UtcNow
    }
}

function Update-SessionTurnState {
    param(
        [string] $Path,
        $Obj,
        [DateTime] $AtUtc
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not $Obj) {
        return
    }

    $type = ""
    try {
        $type = [string]$Obj.type
    } catch {
        $type = ""
    }
    if ($type -ne "event_msg") {
        return
    }

    $eventType = ""
    $turnId = ""
    try {
        $eventType = [string]$Obj.payload.type
        $turnId = [string]$Obj.payload.turn_id
    } catch {
        $eventType = ""
        $turnId = ""
    }
    if ([string]::IsNullOrWhiteSpace($turnId)) {
        return
    }

    Ensure-SessionStatus -Path $Path
    $turns = $script:sessionStatus[$Path].ActiveTurns
    if ($null -eq $turns) {
        $turns = @{}
        $script:sessionStatus[$Path].ActiveTurns = $turns
    }

    switch ($eventType) {
        "task_started" {
            $turns[$turnId] = $AtUtc
            break
        }
        "task_complete" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            break
        }
        "turn_aborted" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            break
        }
        "task_failed" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            break
        }
        "task_cancelled" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            break
        }
    }
}

function Refresh-SessionSnapshot {
    $now = [DateTime]::UtcNow
    if (($now - $script:lastSessionRefreshUtc).TotalMilliseconds -lt 2200) {
        return
    }
    $script:lastSessionRefreshUtc = $now

    $files = @()
    foreach ($dir in (Get-RecentSessionDirs)) {
        try {
            $files += Get-ChildItem -Path $dir -Recurse -File -Filter *.jsonl -ErrorAction SilentlyContinue
        } catch {
        }
    }

    if ($files.Count -eq 0) {
        $script:recentSessionInfos = @()
        return
    }

    $recent = @(
        $files |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 40
    )

    $infos = @()
    foreach ($item in $recent) {
        $path = [string]$item.FullName
        Ensure-SessionStatus -Path $path
        $title = Read-SessionTitle -Item $item
        $infos += [PSCustomObject]@{
            Path = $path
            LastWriteTimeUtc = $item.LastWriteTimeUtc
            FileLength = [int64]$item.Length
            Title = $title
        }
    }
    $script:recentSessionInfos = $infos

    $keep = @{}
    foreach ($info in $infos) {
        $keep[$info.Path] = $true
    }
    foreach ($k in @($script:sessionOffsets.Keys)) {
        if (-not $keep.ContainsKey($k)) {
            $script:sessionOffsets.Remove($k) | Out-Null
        }
    }
    foreach ($k in @($script:sessionStatus.Keys)) {
        if (-not $keep.ContainsKey($k)) {
            $script:sessionStatus.Remove($k) | Out-Null
        }
    }
}

function Is-ApprovalCall {
    param($Payload)
    if (-not $Payload -or $Payload.type -ne "function_call") {
        return $false
    }
    $args = $Payload.arguments
    if ($null -eq $args) {
        return $false
    }
    if ($args -is [string]) {
        return $args -match '"sandbox_permissions"\s*:\s*"require_escalated"'
    }
    $sp = $null
    try {
        $sp = $args.sandbox_permissions
    } catch {
        $sp = $null
    }
    return $sp -eq "require_escalated"
}

function Is-ApprovalMenuLine {
    param([string] $Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $false
    }

    $s = $Line.ToLowerInvariant().Trim()
    if ($s -match '^[›>\s]*1\.\s*yes,\s*proceed\s*\(y\)(\s|$)') {
        return $true
    }
    if ($s -match '^[›>\s]*yes,\s*proceed\s*\(y\)(\s|$)') {
        return $true
    }
    if ($s -match '^[›>\s]*\d+\.\s*yes\b.*\(\s*y\s*\)(\s|$)') {
        return $true
    }

    return $false
}

function Is-PromptishLine {
    param([string] $Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $false
    }

    $s = $Line.ToLowerInvariant()
    if ($s -match '\[y/n\]' -or $s -match '\(y/n\)' -or $s -match '\by/n\b' -or $s -match 'yes/no' -or $s -match '\(\s*y\s*\)' -or $s -match '\[\s*y\s*\]') { return $true }
    if ($s -match 'press\s+enter' -or $s -match 'press\s+any\s+key') { return $true }
    if ($s -match 'press\s+y(es)?' -or $s -match 'type\s+y(es)?' -or $s -match 'enter\s+y(es)?') { return $true }
    if ($s -match 'y\s+to\s+continue' -or $s -match 'y\s+to\s+proceed') { return $true }
    if (Is-ApprovalMenuLine -Line $Line) { return $true }
    return $false
}

function Is-ApprovalTextLikeBark {
    param([string] $Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $false
    }

    $s = $Line.ToLowerInvariant()
    $promptish = Is-PromptishLine -Line $Line
    $hasPolicy = ($s -match 'approval_policy') -or ($s -match 'approval policy')
    $hasApprovalPhrase = ($s -match 'approval required') -or ($s -match 'requires approval') -or ($s -match 'need approval') -or ($s -match 'requesting approval') -or ($s -match 'awaiting approval') -or ($s -match 'waiting for approval')
    $hasProceedPhrase = ($s -match '\bproceed\b') -or ($s -match '\bcontinue\b') -or ($s -match '\bconfirm\b') -or ($s -match 'are you sure') -or ($s -match 'press enter') -or ($s -match 'press any key')
    $hasApprovalWord = ($s -match '\bapprove\b') -or ($s -match '\bapproval\b') -or ($s -match '\bpermission\b') -or ($s -match '\ballow\b') -or ($s -match '\bauthoriz(e|ation)\b') -or ($s -match '\bconsent\b') -or ($s -match '\bproceed\b')
    $hasContextWord = ($s -match 'command') -or ($s -match 'shell') -or ($s -match '\brun\b') -or ($s -match 'execute') -or ($s -match 'sandbox') -or ($s -match 'escalat') -or ($s -match 'network') -or ($s -match 'write') -or ($s -match 'delete') -or ($s -match 'install') -or ($s -match 'access')

    if ($hasApprovalPhrase) { return $true }
    if ($promptish -and $hasPolicy) { return $true }
    if ($promptish -and $hasProceedPhrase -and $hasContextWord) { return $true }
    if ($promptish -and $hasApprovalWord -and $hasContextWord) { return $true }
    return $false
}

function Is-ReconnectTextLikeBark {
    param([string] $Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $false
    }

    $low = $Line.ToLowerInvariant()
    if ($low -match 'reconnecting' -or $low -match '\breconnect\b') {
        return $true
    }

    $flat = ($low -replace '[\s\-]+', '')
    if ($flat -match 'reconnect') {
        return $true
    }

    if ($low -match 'stream disconnected' -or
        $low -match 'retrying turn' -or
        $low -match 'retrying sampling request' -or
        $low -match 'connection lost' -or
        $low -match 'connection closed' -or
        $low -match 'connection dropped') {
        return $true
    }

    $promptish = Is-PromptishLine -Line $Line
    if ($promptish -and
        ($low -match 'disconnect' -or
         $low -match 'disconnected' -or
         $low -match 'connection lost' -or
         $low -match 'connection closed' -or
         $low -match 'connection dropped' -or
         $low -match '\bretry\b' -or
         $low -match 'retrying')) {
        return $true
    }

    return $false
}

function Is-ReconnectRecoveredTextLikeBark {
    param([string] $Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $false
    }

    $low = $Line.ToLowerInvariant()
    if ($low -match 'failed to reconnect' -or
        $low -match 'reconnect failed' -or
        $low -match 'connection lost' -or
        $low -match 'stream disconnected') {
        return $false
    }

    if ($low -match 'connection restored' -or
        $low -match 'reconnected' -or
        $low -match 'reconnect succeeded' -or
        $low -match 'reconnect successful' -or
        $low -match 'reconnect complete' -or
        $low -match 'retry succeeded' -or
        $low -match 'retry successful' -or
        $low -match 'stream resumed' -or
        $low -match 'back online') {
        return $true
    }

    $flat = ($low -replace '[\s\-]+', '')
    if ($flat -match 'connectionrestored' -or
        $flat -match 'reconnectedsuccess' -or
        $flat -match 'reconnectsucceeded' -or
        $flat -match 'streamresumed' -or
        $flat -match 'retrysucceeded') {
        return $true
    }

    return $false
}

function Is-ApprovalSignal {
    param(
        $Obj,
        [string] $Line
    )

    if ($Obj) {
        $type = [string]$Obj.type
        if ($type -eq "response_item") {
            $payload = $Obj.payload
            if ($payload -and (Is-ApprovalCall $payload)) {
                return $true
            }
            try {
                if ([string]$payload.type -eq "approval_request") {
                    return $true
                }
            } catch {
            }
        } elseif ($type -eq "event_msg") {
            $pType = ""
            try {
                $pType = [string]$Obj.payload.type
            } catch {
                $pType = ""
            }
            if ($pType -match "approval|requires_approval") {
                return $true
            }
            try {
                if ($Obj.payload.requires_approval -eq $true) {
                    return $true
                }
            } catch {
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Line)) {
        $s = $Line.ToLowerInvariant()
        $trim = $Line.TrimStart()
        $isJsonRecord = $trim.StartsWith("{")
        if ($s -match '"type"\s*:\s*"response_item"' -and
            $s -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"function_call"' -and
            $s -match '"sandbox_permissions"\s*:\s*"require_escalated"') { return $true }
        if ($s -match '"type"\s*:\s*"event_msg"' -and $s -match '"requires_approval"\s*:\s*true') { return $true }
        if ($s -match '"type"\s*:\s*"event_msg"' -and $s -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"approval_request"') { return $true }
        if (Is-ApprovalMenuLine -Line $Line) { return $true }
        if (-not $isJsonRecord -and (Is-ApprovalTextLikeBark -Line $Line)) { return $true }
    }

    return $false
}

function Is-ReconnectingSignal {
    param(
        $Obj,
        [string] $Line
    )

    if ($Obj) {
        $type = [string]$Obj.type
        if ($type -eq "event_msg") {
            $pType = ""
            try {
                $pType = [string]$Obj.payload.type
            } catch {
                $pType = ""
            }
            if ($pType -match "reconnect|reconnecting|connection_lost") {
                return $true
            }
            $msg = ""
            try {
                $msg = [string]$Obj.payload.message
            } catch {
                $msg = ""
            }
            if (-not [string]::IsNullOrWhiteSpace($msg) -and (Is-ReconnectRecoveredTextLikeBark -Line $msg)) {
                return $false
            }
            if (-not [string]::IsNullOrWhiteSpace($msg) -and (Is-ReconnectTextLikeBark -Line $msg)) {
                return $true
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Line)) {
        $s = $Line.ToLowerInvariant()
        $trim = $Line.TrimStart()
        $isJsonRecord = $trim.StartsWith("{")
        if (Is-ReconnectRecoveredTextLikeBark -Line $Line) { return $false }
        if ($isJsonRecord) {
            if ($s -match '"type"\s*:\s*"event_msg"' -and
                $s -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"(reconnect|reconnecting|connection_lost)"') { return $true }
            if ($s -match '"type"\s*:\s*"event_msg"' -and
                ($s -match 'stream disconnected' -or $s -match 'retrying turn' -or $s -match 'retrying sampling request' -or $s -match 'connection lost' -or $s -match 'connection closed' -or $s -match 'connection dropped')) { return $true }
        } else {
            if (Is-ReconnectTextLikeBark -Line $Line) { return $true }
        }
    }

    return $false
}

function Is-ReconnectRecoveredSignal {
    param(
        $Obj,
        [string] $Line
    )

    if ($Obj) {
        $type = [string]$Obj.type
        if ($type -eq "event_msg") {
            $pType = ""
            try {
                $pType = [string]$Obj.payload.type
            } catch {
                $pType = ""
            }
            if ($pType -match "connection_restored|reconnected|reconnect_succeeded|reconnect_successful|stream_resumed|retry_succeeded") {
                return $true
            }
            $msg = ""
            try {
                $msg = [string]$Obj.payload.message
            } catch {
                $msg = ""
            }
            if (-not [string]::IsNullOrWhiteSpace($msg) -and (Is-ReconnectRecoveredTextLikeBark -Line $msg)) {
                return $true
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Line)) {
        $s = $Line.ToLowerInvariant()
        $trim = $Line.TrimStart()
        $isJsonRecord = $trim.StartsWith("{")
        if ($isJsonRecord) {
            if ($s -match '"type"\s*:\s*"event_msg"' -and
                $s -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"(connection_restored|reconnected|reconnect_succeeded|reconnect_successful|stream_resumed|retry_succeeded)"') { return $true }
        }
        if (Is-ReconnectRecoveredTextLikeBark -Line $Line) { return $true }
    }

    return $false
}

function Is-ReconnectClearActivity {
    param(
        $Obj,
        [string] $Line
    )

    if (-not [string]::IsNullOrWhiteSpace($Line)) {
        if ($Line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"(function_call_output|custom_tool_call_output)"') {
            return $true
        }
    }

    if (-not $Obj) {
        return $false
    }

    $type = [string]$Obj.type
    if ($type -eq "response_item") {
        $payloadType = ""
        $role = ""
        try { $payloadType = [string]$Obj.payload.type } catch { $payloadType = "" }
        try { $role = [string]$Obj.payload.role } catch { $role = "" }
        if ($payloadType -eq "message" -and $role -eq "assistant") {
            return $true
        }
        if ($payloadType -eq "function_call_output" -or $payloadType -eq "custom_tool_call_output") {
            return $true
        }
    }

    if ($type -eq "event_msg") {
        $eventType = ""
        try { $eventType = [string]$Obj.payload.type } catch { $eventType = "" }
        if ($eventType -match "task_started|task_complete|turn_aborted|task_failed|task_cancelled|agent_message|approval_request|requires_approval") {
            return $true
        }
    }

    return $false
}

function Update-SessionTurnStateFromLine {
    param(
        [string] $Path,
        [string] $Line,
        [DateTime] $AtUtc
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Line)) {
        return
    }
    if ($Line -notmatch '"type"\s*:\s*"event_msg"') {
        return
    }

    $eventType = ""
    $turnId = ""
    $mType = [regex]::Match($Line, '"payload"\s*:\s*\{\s*"type"\s*:\s*"([^"]+)"')
    if ($mType.Success) {
        $eventType = [string]$mType.Groups[1].Value
    }
    $mTurn = [regex]::Match($Line, '"turn_id"\s*:\s*"([^"]+)"')
    if ($mTurn.Success) {
        $turnId = [string]$mTurn.Groups[1].Value
    }

    Ensure-SessionStatus -Path $Path
    $turns = $script:sessionStatus[$Path].ActiveTurns
    if ($null -eq $turns) {
        $turns = @{}
        $script:sessionStatus[$Path].ActiveTurns = $turns
    }

    if ([string]::IsNullOrWhiteSpace($turnId)) {
        if ($eventType -eq "task_complete" -or $eventType -eq "turn_aborted" -or $eventType -eq "task_failed" -or $eventType -eq "task_cancelled") {
            Clear-SessionApprovalPending -Path $Path
            Clear-SessionReconnectPending -Path $Path
        }
        return
    }

    switch ($eventType) {
        "task_started" {
            $turns[$turnId] = $AtUtc
            Clear-SessionReconnectPending -Path $Path
            break
        }
        "task_complete" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            Clear-SessionApprovalPending -Path $Path
            Clear-SessionReconnectPending -Path $Path
            break
        }
        "turn_aborted" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            Clear-SessionApprovalPending -Path $Path
            Clear-SessionReconnectPending -Path $Path
            break
        }
        "task_failed" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            Clear-SessionApprovalPending -Path $Path
            Clear-SessionReconnectPending -Path $Path
            break
        }
        "task_cancelled" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            Clear-SessionApprovalPending -Path $Path
            Clear-SessionReconnectPending -Path $Path
            break
        }
    }
}

function Parse-SessionLine {
    param(
        [string] $Path,
        [string] $Line
    )
    if ([string]::IsNullOrWhiteSpace($Line)) {
        return
    }

    $eventUtc = Get-LineTimestampUtc -Line $Line
    $speech = Extract-AssistantSpeechFromLine -Line $Line
    if (-not [string]::IsNullOrWhiteSpace($speech)) {
        Mark-SessionSpeech -Path $Path -Speech $speech -AtUtc $eventUtc
    }
    $isEventMsg = ($Line -match '"type"\s*:\s*"event_msg"')
    $isResponseItem = ($Line -match '"type"\s*:\s*"response_item"')
    $isFunctionCall = ($Line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"function_call"')
    $isFunctionCallOutput = ($Line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"(function_call_output|custom_tool_call_output)"')
    $obj = $null
    if ($isEventMsg -or $isResponseItem) {
        try {
            $obj = $Line | ConvertFrom-Json -ErrorAction Stop
        } catch {
            $obj = $null
        }
    }

    $approvalFromLine = Is-ApprovalSignal -Obj $obj -Line $Line
    $hasApprovalHint = $approvalFromLine
    $reconnectRecoveredFromLine = Is-ReconnectRecoveredSignal -Obj $obj -Line $Line
    $reconnectFromLine = $false
    if (-not $reconnectRecoveredFromLine) {
        $reconnectFromLine = Is-ReconnectingSignal -Obj $obj -Line $Line
    }
    $isFinalAnswer = Is-FinalAnswerSignal -Line $Line
    $isTerminalActivity = Is-TerminalActivitySignal -Obj $obj -Line $Line
    $isWorkActivity = Is-WorkActivitySignal -Obj $obj

    # Approval should clear only when we see real resumed assistant output after the approval event.
    if ($script:sessionStatus.ContainsKey($Path)) {
        $st = $script:sessionStatus[$Path]
        $isAfterApproval = $false
        try {
            $isAfterApproval = ($st.LastApprovalUtc -ne [DateTime]::MinValue -and $eventUtc -gt $st.LastApprovalUtc)
        } catch {
            $isAfterApproval = $false
        }
        if ($st.ApprovalPending -and $isAfterApproval) {
            $assistantOutput = $false
            if ($obj) {
                try {
                    if ([string]$obj.type -eq "response_item" -and
                        [string]$obj.payload.type -eq "message" -and
                        [string]$obj.payload.role -eq "assistant") {
                        $assistantOutput = $true
                    }
                } catch {
                }
                try {
                    if ([string]$obj.type -eq "event_msg" -and
                        [string]$obj.payload.type -eq "agent_message") {
                        $assistantOutput = $true
                    }
                } catch {
                }
            }
            if ($assistantOutput) {
                Clear-SessionApprovalPending -Path $Path
            }
        }

        $isAfterReconnect = $false
        try {
            $isAfterReconnect = ($st.LastReconnectUtc -ne [DateTime]::MinValue -and $eventUtc -gt $st.LastReconnectUtc)
        } catch {
            $isAfterReconnect = $false
        }
        if ($st.ReconnectPending -and $isAfterReconnect) {
            $clearReconnect = $reconnectRecoveredFromLine
            if (-not $clearReconnect -and -not $reconnectFromLine) {
                $clearReconnect = Is-ReconnectClearActivity -Obj $obj -Line $Line
            }
            if ($clearReconnect) {
                Clear-SessionReconnectPending -Path $Path
            }
        }
    }

    if ($isFunctionCallOutput) {
        Clear-SessionApprovalPending -Path $Path
        Clear-SessionReconnectPending -Path $Path
    }
    if ($isTerminalActivity) {
        Mark-SessionTerminal -Path $Path -AtUtc $eventUtc
    }
    if ($isFinalAnswer) {
        Clear-SessionActiveTurns -Path $Path
        Clear-SessionApprovalPending -Path $Path
        Clear-SessionReconnectPending -Path $Path
    } elseif ($isWorkActivity) {
        Mark-SessionWork -Path $Path -AtUtc $eventUtc
    }

    if (-not $isEventMsg -and -not $isResponseItem) {
        if ($reconnectRecoveredFromLine) {
            Clear-SessionReconnectPending -Path $Path
            Mark-SessionEvent -Path $Path -AtUtc $eventUtc
            return
        }
        if ($reconnectFromLine) {
            Mark-SessionEvent -Path $Path -Reconnecting -AtUtc $eventUtc
            return
        }
        if ($approvalFromLine) {
            Mark-SessionEvent -Path $Path -Approval -AtUtc $eventUtc
            return
        }
        return
    }

    if ($isEventMsg) {
        Update-SessionTurnStateFromLine -Path $Path -Line $Line -AtUtc $eventUtc

        if ($approvalFromLine) {
            Mark-SessionEvent -Path $Path -Approval -AtUtc $eventUtc
            return
        }

        if ($reconnectRecoveredFromLine) {
            Clear-SessionReconnectPending -Path $Path
            Mark-SessionEvent -Path $Path -AtUtc $eventUtc
            return
        }

        if ($reconnectFromLine) {
            Mark-SessionEvent -Path $Path -Reconnecting -AtUtc $eventUtc
            return
        }

        Mark-SessionEvent -Path $Path -AtUtc $eventUtc
        return
    }

    if ($isFunctionCall -or $hasApprovalHint) {
        if ($approvalFromLine) {
            Mark-SessionEvent -Path $Path -Approval -AtUtc $eventUtc
            return
        }
    }

    if ($reconnectRecoveredFromLine) {
        Clear-SessionReconnectPending -Path $Path
    }
    if ($reconnectFromLine) {
        Mark-SessionEvent -Path $Path -Reconnecting -AtUtc $eventUtc
        return
    }

    Mark-SessionEvent -Path $Path -AtUtc $eventUtc
}

function Poll-SessionFile {
    param(
        [string] $Path,
        [int] $LineBudget = 80
    )
    if ($LineBudget -le 0 -or [string]::IsNullOrWhiteSpace($Path)) {
        return 0
    }

    $item = Get-Item -Path $Path -ErrorAction SilentlyContinue
    if (-not $item) {
        $script:sessionOffsets.Remove($Path) | Out-Null
        return 0
    }

    if (-not $script:sessionOffsets.ContainsKey($Path)) {
        Ensure-SessionStatus -Path $Path
        $ageSec = ([DateTime]::UtcNow - $item.LastWriteTimeUtc).TotalSeconds
        if ($ageSec -le 14400.0) {
            Initialize-SessionFromTail -Path $Path -Length ([int64]$item.Length)
        } else {
            $script:sessionStatus[$Path].ActiveTurns = @{}
        }
        $script:sessionOffsets[$Path] = [int64]$item.Length
        return 0
    }

    $length = [int64]$item.Length
    $offset = [int64]$script:sessionOffsets[$Path]

    if ($length -lt $offset) {
        Ensure-SessionStatus -Path $Path
        $script:sessionStatus[$Path].ActiveTurns = @{}
        Initialize-SessionFromTail -Path $Path -Length $length
        $offset = [int64]$length
        $script:sessionOffsets[$Path] = $offset
        return 0
    }
    if ($length -eq $offset) {
        return 0
    }

    $fs = $null
    $sr = $null
    $processed = 0
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        [void]$fs.Seek($offset, [System.IO.SeekOrigin]::Begin)
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true, 4096, $true)
        while ((-not $sr.EndOfStream) -and $processed -lt $LineBudget) {
            $line = $sr.ReadLine()
            Parse-SessionLine -Path $Path -Line $line
            $processed++
        }
        if ($fs.Position -gt $offset) {
            Mark-SessionEvent -Path $Path
        }
        $script:sessionOffsets[$Path] = $fs.Position
    } catch {
    } finally {
        if ($sr) { $sr.Dispose() }
        if ($fs) { $fs.Dispose() }
    }

    return $processed
}

function Initialize-SessionFromTail {
    param(
        [string] $Path,
        [int64] $Length
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or $Length -le 0) {
        return
    }

    # bootstrap active turn state from recent tail (fast startup)
    $bootstrapBytes = [int64]8388608
    $start = [Math]::Max([int64]0, $Length - $bootstrapBytes)

    Ensure-SessionStatus -Path $Path
    $turns = @{}
    $latestEventUtc = [DateTime]::MinValue

    $fs = $null
    $sr = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        [void]$fs.Seek($start, [System.IO.SeekOrigin]::Begin)
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true, 4096, $true)
        if ($start -gt 0 -and -not $sr.EndOfStream) {
            [void]$sr.ReadLine()
        }
        while (-not $sr.EndOfStream) {
            $line = $sr.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            $tsMatch = [regex]::Match($line, '"timestamp"\s*:\s*"([^"]+)"')
            $eventUtc = [DateTime]::UtcNow
            if ($tsMatch.Success) {
                try {
                    $eventUtc = ([DateTimeOffset]::Parse($tsMatch.Groups[1].Value)).UtcDateTime
                } catch {
                    $eventUtc = [DateTime]::UtcNow
                }
            }
            if ($eventUtc -gt $latestEventUtc) {
                $latestEventUtc = $eventUtc
            }

            $speech = Extract-AssistantSpeechFromLine -Line $line
            if (-not [string]::IsNullOrWhiteSpace($speech)) {
                Mark-SessionSpeech -Path $Path -Speech $speech -AtUtc $eventUtc
            }

            if (Is-FinalAnswerSignal -Line $line) {
                Mark-SessionTerminal -Path $Path -AtUtc $eventUtc
                $turns = @{}
                $script:sessionStatus[$Path].ActiveTurns = $turns
                Clear-SessionApprovalPending -Path $Path
                Clear-SessionReconnectPending -Path $Path
            }
            if ($line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"(function_call_output|custom_tool_call_output)"') {
                Clear-SessionApprovalPending -Path $Path
                Clear-SessionReconnectPending -Path $Path
            }
            $obj = $null
            if ($line -match '"type"\s*:\s*"(event_msg|response_item)"') {
                try {
                    $obj = $line | ConvertFrom-Json -ErrorAction Stop
                } catch {
                    $obj = $null
                }
            }
            if (Is-TerminalActivitySignal -Obj $obj -Line $line) {
                Mark-SessionTerminal -Path $Path -AtUtc $eventUtc
            } elseif (Is-WorkActivitySignal -Obj $obj) {
                Mark-SessionWork -Path $Path -AtUtc $eventUtc
            }
            if (Is-ApprovalSignal -Obj $obj -Line $line) {
                Mark-SessionEvent -Path $Path -Approval -AtUtc $eventUtc
            }
            $isReconnectRecovered = Is-ReconnectRecoveredSignal -Obj $obj -Line $line
            $isReconnect = $false
            if (-not $isReconnectRecovered) {
                $isReconnect = Is-ReconnectingSignal -Obj $obj -Line $line
            }
            if ($isReconnectRecovered) {
                Clear-SessionReconnectPending -Path $Path
            } elseif ($isReconnect) {
                Mark-SessionEvent -Path $Path -Reconnecting -AtUtc $eventUtc
            } elseif (Is-ReconnectClearActivity -Obj $obj -Line $line) {
                Clear-SessionReconnectPending -Path $Path
            }

            if ($line.IndexOf('"type":"event_msg"', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                continue
            }
            if ($line.IndexOf('"turn_id"', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                continue
            }

            $tidMatch = [regex]::Match($line, '"turn_id"\s*:\s*"([^"]+)"')
            if (-not $tidMatch.Success) {
                continue
            }
            $turnId = $tidMatch.Groups[1].Value
            if ([string]::IsNullOrWhiteSpace($turnId)) {
                continue
            }

            if ($line.IndexOf('"payload":{"type":"task_started"', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"task_started"') {
                $turns[$turnId] = $eventUtc
                Clear-SessionReconnectPending -Path $Path
                continue
            }

            if ($line.IndexOf('"payload":{"type":"task_complete"', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $line.IndexOf('"payload":{"type":"turn_aborted"', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $line.IndexOf('"payload":{"type":"task_failed"', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $line.IndexOf('"payload":{"type":"task_cancelled"', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"(task_complete|turn_aborted|task_failed|task_cancelled)"') {
                if ($turns.ContainsKey($turnId)) {
                    $turns.Remove($turnId) | Out-Null
                }
                Clear-SessionApprovalPending -Path $Path
                Clear-SessionReconnectPending -Path $Path
                continue
            }
        }
    } catch {
    } finally {
        if ($sr) { $sr.Dispose() }
        if ($fs) { $fs.Dispose() }
    }

    $script:sessionStatus[$Path].ActiveTurns = $turns
    if ($latestEventUtc -gt $script:sessionStatus[$Path].LastEventUtc) {
        $script:sessionStatus[$Path].LastEventUtc = $latestEventUtc
    }
}

function Poll-Sessions {
    Refresh-SessionSnapshot
    Refresh-ReconnectSignalsFromCodexLogs
    if ($script:recentSessionInfos.Count -eq 0) {
        return
    }

    $remaining = 520
    foreach ($info in @($script:recentSessionInfos | Select-Object -First 12)) {
        if ($remaining -le 0) {
            break
        }
        $used = Poll-SessionFile -Path ([string]$info.Path) -LineBudget $remaining
        if ($used -gt 0) {
            $remaining -= $used
        }
    }
}

function Get-CodexAgentCount {
    $now = [DateTime]::UtcNow
    if (($now - $script:lastCodexProbeUtc).TotalMilliseconds -lt 900) {
        return $script:lastCodexCount
    }

    Refresh-ActiveRolloutPathSet
    $activeCount = 0
    try {
        $activeCount = [int]$script:activeRolloutPathSet.Count
    } catch {
        $activeCount = 0
    }

    $processCount = 0
    try {
        $processCount = @(Get-Process -Name codex -ErrorAction SilentlyContinue).Count
    } catch {
        $processCount = 0
    }

    $count = [Math]::Max($activeCount, $processCount)
    if ($count -eq 0 -and $script:currentState -ne "idle") {
        $count = 1
    }

    $script:lastCodexProbeUtc = $now
    $script:lastCodexCount = $count
    return $count
}

function Get-SessionStateInfo {
    param(
        [string] $Path,
        [DateTime] $LastWriteTimeUtc,
        [int64] $FileLength = 0
    )

    $now = [DateTime]::UtcNow
    $approvalHoldSec = [Math]::Max(90.0, $IdleSeconds * 20.0)
    $approvalPendingMaxSec = [Math]::Max(21600.0, $IdleSeconds * 14400.0)
    $approvalActiveTurnGuardSec = [Math]::Max(300.0, $IdleSeconds * 300.0)
    $reconnectPendingMaxSec = [Math]::Max(43200.0, $IdleSeconds * 21600.0)
    $reconnectRecentHoldSec = [Math]::Max(90.0, $IdleSeconds * 45.0)
    $runningWindowSec = [Math]::Max(30.0, $IdleSeconds * 20.0)
    $stickyWorkMaxSilenceSec = [Math]::Max(7200.0, $IdleSeconds * 4800.0)
    $status = $null
    if ($script:sessionStatus.ContainsKey($Path)) {
        $status = $script:sessionStatus[$Path]
    }

    $hasActiveTurn = $false
    $hasStickyWork = $false
    $lastWriteUtc = ([DateTime]$LastWriteTimeUtc).ToUniversalTime()
    $writeAgeSec = ($now - $lastWriteUtc).TotalSeconds
    $activeTurnMaxSilenceSec = [Math]::Max(7200.0, $IdleSeconds * 4800.0)
    if ($status -and $null -ne $status.ActiveTurns -and $status.ActiveTurns.Count -gt 0) {
        $cutoffUtc = [DateTime]::MinValue
        if ($status.LastTerminalUtc -ne [DateTime]::MinValue) {
            $cutoffUtc = $status.LastTerminalUtc
        }
        $keptTurns = @{}
        foreach ($k in @($status.ActiveTurns.Keys)) {
            $dt = [DateTime]$status.ActiveTurns[$k]
            if ($cutoffUtc -ne [DateTime]::MinValue -and $dt -le $cutoffUtc) {
                continue
            }
            $keptTurns[$k] = $dt
        }
        $status.ActiveTurns = $keptTurns
        if ($status.ActiveTurns.Count -gt 0) {
            if ($writeAgeSec -le $activeTurnMaxSilenceSec) {
                $hasActiveTurn = $true
            } else {
                $status.ActiveTurns = @{}
            }
        }
    }

    if ($status -and $status.LastWorkUtc -ne [DateTime]::MinValue) {
        $workStillOpen = ($status.LastTerminalUtc -eq [DateTime]::MinValue -or $status.LastWorkUtc -gt $status.LastTerminalUtc)
        if ($workStillOpen) {
            $workSignalUtc = $status.LastWorkUtc
            if ($lastWriteUtc -gt $workSignalUtc) {
                $workSignalUtc = $lastWriteUtc
            }
            $workAgeSec = ($now - $workSignalUtc).TotalSeconds
            if ($workAgeSec -le $stickyWorkMaxSilenceSec) {
                $hasStickyWork = $true
            }
        }
    }

    $lastSignal = $lastWriteUtc
    if ($status -and $status.LastEventUtc -gt $lastSignal) {
        $lastSignal = $status.LastEventUtc
    }
    $age = ($now - $lastSignal).TotalSeconds
    if ($age -lt 0) {
        $age = 0
    }

    if ($status -and $status.ApprovalPending) {
        $approvalAge = ($now - $status.LastApprovalUtc).TotalSeconds
        if ($writeAgeSec -le $approvalPendingMaxSec) {
            return [PSCustomObject]@{
                State = "approval"
                AgeSec = [int][Math]::Floor([Math]::Max(0.0, $approvalAge))
            }
        } else {
            $status.ApprovalPending = $false
        }
    }

    if ($status -and $status.ReconnectPending) {
        $reconnectAge = ($now - $status.LastReconnectUtc).TotalSeconds
        if ($writeAgeSec -le $reconnectPendingMaxSec) {
            return [PSCustomObject]@{
                State = "reconnecting"
                AgeSec = [int][Math]::Floor([Math]::Max(0.0, $reconnectAge))
            }
        } else {
            $status.ReconnectPending = $false
        }
    }

    if ($status -and $status.LastReconnectUtc -ne [DateTime]::MinValue) {
        $reconnectAge = ($now - $status.LastReconnectUtc).TotalSeconds
        if ($reconnectAge -le $reconnectRecentHoldSec) {
            return [PSCustomObject]@{
                State = "reconnecting"
                AgeSec = [int][Math]::Floor([Math]::Max(0.0, $reconnectAge))
            }
        }
    }

    $hasOngoingWork = ($hasActiveTurn -or $hasStickyWork)

    if ($status -and $status.LastApprovalUtc -ne [DateTime]::MinValue -and -not $hasOngoingWork) {
        $approvalAge = ($now - $status.LastApprovalUtc).TotalSeconds
        if ($approvalAge -le $approvalHoldSec) {
            return [PSCustomObject]@{
                State = "approval"
                AgeSec = [int][Math]::Floor([Math]::Max(0.0, $approvalAge))
            }
        }
    }

    if ($hasOngoingWork) {
        if ($status -and $status.LastApprovalUtc -ne [DateTime]::MinValue) {
            $approvalAgeForActive = ($now - $status.LastApprovalUtc).TotalSeconds
            if ($approvalAgeForActive -le $approvalActiveTurnGuardSec -and $age -gt $runningWindowSec) {
                return [PSCustomObject]@{
                    State = "approval"
                    AgeSec = [int][Math]::Floor([Math]::Max(0.0, $approvalAgeForActive))
                }
            }
        }
        if ($age -le $runningWindowSec) {
            return [PSCustomObject]@{
                State = "running"
                AgeSec = [int][Math]::Floor([Math]::Max(0.0, $age))
            }
        }
        return [PSCustomObject]@{
            State = "silent"
            AgeSec = [int][Math]::Floor([Math]::Max(0.0, $age))
        }
    }

    if ($age -le $runningWindowSec) {
        return [PSCustomObject]@{
            State = "running"
            AgeSec = [int][Math]::Floor([Math]::Max(0.0, $age))
        }
    }

    return [PSCustomObject]@{
        State = "idle"
        AgeSec = [int][Math]::Floor([Math]::Max(0.0, $age))
    }
}

function Get-CardModels {
    param([int] $Count)
    if ($Count -le 0) {
        return @()
    }

    Refresh-SessionSnapshot
    Refresh-ActiveRolloutPathSet
    $nowUtc = [DateTime]::UtcNow
    $activePathSet = @{}
    foreach ($k in @($script:activeRolloutPathSet.Keys)) {
        $activePathSet[[string]$k] = $true
    }

    $bootstrapVisibleWindowSec = 45.0
    $candidateInfos = @(
        $script:recentSessionInfos |
            Where-Object {
                $p = [string]$_.Path
                $np = Normalize-SessionPath -Path $p
                if (-not [string]::IsNullOrWhiteSpace($np) -and $activePathSet.ContainsKey($np)) {
                    return $true
                }
                if ($script:sessionStatus.ContainsKey($p)) {
                    $st = $script:sessionStatus[$p]
                    if ($st) {
                        if ($null -ne $st.ActiveTurns -and $st.ActiveTurns.Count -gt 0) {
                            return $true
                        }
                        if ($st.ApprovalPending -or $st.ReconnectPending) {
                            return $true
                        }
                        if ($st.LastWorkUtc -ne [DateTime]::MinValue -and
                            ($st.LastTerminalUtc -eq [DateTime]::MinValue -or $st.LastWorkUtc -gt $st.LastTerminalUtc)) {
                            return $true
                        }
                    }
                }
                $ageSec = ($nowUtc - ([DateTime]$_.LastWriteTimeUtc).ToUniversalTime()).TotalSeconds
                if ($ageSec -le $bootstrapVisibleWindowSec) {
                    return $true
                }
                return $false
            }
    )

    if ($activePathSet.Count -gt 0) {
        $have = @{}
        foreach ($info in $candidateInfos) {
            $np = Normalize-SessionPath -Path ([string]$info.Path)
            if (-not [string]::IsNullOrWhiteSpace($np)) {
                $have[$np] = $true
            }
        }
        foreach ($np in @($activePathSet.Keys)) {
            if ($have.ContainsKey($np)) {
                continue
            }
            $rawPath = [string]$np
            if (-not (Test-Path $rawPath)) {
                continue
            }
            try {
                $item = Get-Item -Path $rawPath -ErrorAction Stop
                if (-not $item -or $item.PSIsContainer -or $item.Extension -ne ".jsonl") {
                    continue
                }
                Ensure-SessionStatus -Path ([string]$item.FullName)
                $candidateInfos += [PSCustomObject]@{
                    Path = [string]$item.FullName
                    LastWriteTimeUtc = $item.LastWriteTimeUtc
                    FileLength = [int64]$item.Length
                    Title = (Read-SessionTitle -Item $item)
                }
                $have[$np] = $true
            } catch {
            }
        }
    }

    if ($candidateInfos.Count -eq 0) {
        $candidateInfos = @($script:recentSessionInfos | Select-Object -First 12)
    }

    $orderedInfos = @(
        $candidateInfos |
            Sort-Object `
                @{ Expression = {
                        $p = Normalize-SessionPath -Path ([string]$_.Path)
                        if (-not [string]::IsNullOrWhiteSpace($p) -and $activePathSet.ContainsKey($p)) {
                            0
                        } else {
                            1
                        }
                    }
                }, `
                @{ Expression = {
                        $p = [string]$_.Path
                        $isFresh = (($nowUtc - ([DateTime]$_.LastWriteTimeUtc).ToUniversalTime()).TotalSeconds -le 120.0)
                        if ($isFresh -and $script:sessionStatus.ContainsKey($p) -and $null -ne $script:sessionStatus[$p].ActiveTurns -and $script:sessionStatus[$p].ActiveTurns.Count -gt 0) {
                            0
                        } else {
                            1
                        }
                    }
                }, `
                @{ Expression = { $_.LastWriteTimeUtc }; Descending = $true }
    )

    $picked = @($orderedInfos | Select-Object -First $Count)
    $selected = @($picked | Select-Object -First $Count)
    $rawModels = @()
    foreach ($info in $selected) {
        $title = [string]$info.Title
        if ([string]::IsNullOrWhiteSpace($title)) {
            $title = "codex"
        }
        $path = [string]$info.Path
        $stateInfo = Get-SessionStateInfo -Path ([string]$info.Path) -LastWriteTimeUtc ([DateTime]$info.LastWriteTimeUtc) -FileLength ([int64]$info.FileLength)
        $rawModels += [PSCustomObject]@{
            Title = $title
            State = [string]$stateInfo.State
            AgeSec = [int]$stateInfo.AgeSec
            Speech = (Get-SessionSpeechPreview -Path $path)
            Key = $path
        }
    }

    while ($rawModels.Count -lt $Count) {
        $rawModels += [PSCustomObject]@{
            Title = "codex"
            State = "idle"
            AgeSec = 0
            Speech = "..."
            Key = ("__empty_{0}" -f $rawModels.Count)
        }
    }

    $seen = @{}
    $result = @()
    foreach ($m in $rawModels) {
        $t = [string]$m.Title
        $k = $t.ToLowerInvariant()
        if ($seen.ContainsKey($k)) {
            $seen[$k] = [int]$seen[$k] + 1
            $display = "{0} ({1})" -f $t, $seen[$k]
        } else {
            $seen[$k] = 1
            $display = $t
        }
        $result += [PSCustomObject]@{
            Title = $display
            State = [string]$m.State
            AgeSec = [int]$m.AgeSec
            Speech = [string]$m.Speech
            Key = [string]$m.Key
        }
    }

    $ordered = Apply-SlotOrder -Models $result -Count $Count
    $final = @()
    for ($i = 0; $i -lt $ordered.Count; $i++) {
        $m = $ordered[$i]
        $final += [PSCustomObject]@{
            Title = [string]$m.Title
            State = [string]$m.State
            AgeSec = [int]$m.AgeSec
            Speech = [string]$m.Speech
            Key = [string]$m.Key
            HotkeyIndex = [int]($i + 1)
        }
    }
    return @($final)
}

function Get-StateText {
    param([string] $State)
    switch ($State) {
        "running" { return "working" }
        "approval" { return "approval" }
        "reconnecting" { return "reconnecting" }
        "silent" { return "silent" }
        default { return "idle" }
    }
}

Ensure-ParentDir $StateFile
if (-not (Test-Path $StateFile)) {
    Write-State "idle"
    $script:currentState = "idle"
}
Load-TabMap
Load-SlotOrder
Load-AnchorOffset

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public struct ConfirwaRect {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}
public static class ConfirwaNative {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out ConfirwaRect lpRect);
    [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
}
"@ -ErrorAction SilentlyContinue
Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue

$script:cards = @{}

function Is-ConfirwaHandle {
    param([IntPtr] $Hwnd)
    if ($Hwnd -eq [IntPtr]::Zero) {
        return $false
    }

    if ($script:driverForm -and $script:driverForm.IsHandleCreated -and $script:driverForm.Handle -eq $Hwnd) {
        return $true
    }

    foreach ($c in @($script:cards.Values)) {
        if ($c -and $c.Form -and $c.Form.IsHandleCreated -and $c.Form.Handle -eq $Hwnd) {
            return $true
        }
    }
    return $false
}

function Is-TerminalProcessName {
    param([string] $Name)
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $false
    }
    switch ($Name.ToLowerInvariant()) {
        "windowsterminal" { return $true }
        "wezterm-gui" { return $true }
        "tabby" { return $true }
        "conemu" { return $true }
        "conemu64" { return $true }
        "pwsh" { return $true }
        "powershell" { return $true }
        "cmd" { return $true }
        default { return $false }
    }
}

function Get-ProcessNameByWindowHandle {
    param([IntPtr] $Hwnd)
    if ($Hwnd -eq [IntPtr]::Zero) {
        return ""
    }
    $procId = [uint32]0
    try {
        [void][ConfirwaNative]::GetWindowThreadProcessId($Hwnd, [ref]$procId)
        if ($procId -eq 0) {
            return ""
        }
        $p = Get-Process -Id ([int]$procId) -ErrorAction SilentlyContinue
        if ($p) {
            return [string]$p.ProcessName
        }
    } catch {
    }
    return ""
}

function Get-ProcessIdByWindowHandle {
    param([IntPtr] $Hwnd)
    if ($Hwnd -eq [IntPtr]::Zero) {
        return 0
    }
    $procId = [uint32]0
    try {
        [void][ConfirwaNative]::GetWindowThreadProcessId($Hwnd, [ref]$procId)
    } catch {
        $procId = 0
    }
    return [int]$procId
}

function Refresh-LastTerminalWindowFromForeground {
    try {
        $fg = [ConfirwaNative]::GetForegroundWindow()
        if ($fg -eq [IntPtr]::Zero) {
            return
        }
        if (Is-ConfirwaHandle -Hwnd $fg) {
            return
        }
        $name = Get-ProcessNameByWindowHandle -Hwnd $fg
        if (Is-TerminalProcessName -Name $name) {
            $script:lastTerminalHwnd = $fg
        }
    } catch {
    }
}

function Find-TerminalWindowHandle {
    $priorityNames = @("WindowsTerminal", "wezterm-gui", "tabby", "ConEmu64", "ConEmu", "pwsh", "powershell", "cmd")
    foreach ($name in $priorityNames) {
        $procs = @(
            Get-Process -Name $name -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne 0 } |
                Sort-Object StartTime -Descending
        )
        foreach ($p in $procs) {
            $h = [IntPtr]$p.MainWindowHandle
            if ($h -eq [IntPtr]::Zero) {
                continue
            }
            if (Is-ConfirwaHandle -Hwnd $h) {
                continue
            }
            return $h
        }
    }
    return [IntPtr]::Zero
}

function Focus-TerminalWindow {
    $target = [IntPtr]$script:lastTerminalHwnd
    if ($target -eq [IntPtr]::Zero -or (Is-ConfirwaHandle -Hwnd $target)) {
        $target = Find-TerminalWindowHandle
    }
    if ($target -eq [IntPtr]::Zero) {
        return $false
    }

    try {
        $targetProcId = Get-ProcessIdByWindowHandle -Hwnd $target
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            if ([ConfirwaNative]::IsIconic($target)) {
                [void][ConfirwaNative]::ShowWindowAsync($target, 9)
            } else {
                [void][ConfirwaNative]::ShowWindowAsync($target, 5)
            }
            Start-Sleep -Milliseconds 35
            [void][ConfirwaNative]::SetForegroundWindow($target)
            Start-Sleep -Milliseconds 35

            $fg = [ConfirwaNative]::GetForegroundWindow()
            if ($fg -eq $target) {
                $script:lastTerminalHwnd = $target
                return $true
            }

            if ($targetProcId -gt 0) {
                try {
                    [void][Microsoft.VisualBasic.Interaction]::AppActivate($targetProcId)
                    Start-Sleep -Milliseconds 45
                    $fg2 = [ConfirwaNative]::GetForegroundWindow()
                    if ($fg2 -eq $target) {
                        $script:lastTerminalHwnd = $target
                        return $true
                    }
                } catch {
                }
            }
        }
        Write-ConfirwaError ("focus failed: target=0x{0:X}" -f ([int64]$target))
        return $false
    } catch {
        Write-ConfirwaError ("focus exception: target=0x{0:X}, err={1}" -f ([int64]$target), ([string]$_.Exception))
        return $false
    }
}

function Send-TerminalTabHotkey {
    param([int] $TabIndex)
    if ($TabIndex -le 0) {
        return
    }
    $digit = $TabIndex
    if ($digit -gt 9) {
        $digit = (($digit - 1) % 9) + 1
    }

    if (-not (Focus-TerminalWindow)) {
        Write-ConfirwaError ("tab switch skipped: no terminal window for card {0}" -f $TabIndex)
        return
    }

    try {
        $fgBefore = [ConfirwaNative]::GetForegroundWindow()
        $vkCtrl = [byte]0x11
        $vkAlt = [byte]0x12
        $vkDigit = [byte](0x30 + $digit)
        $keyUp = [uint32]0x0002

        [ConfirwaNative]::keybd_event($vkCtrl, 0, 0, [UIntPtr]::Zero)
        [ConfirwaNative]::keybd_event($vkAlt, 0, 0, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 10
        [ConfirwaNative]::keybd_event($vkDigit, 0, 0, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 16
        [ConfirwaNative]::keybd_event($vkDigit, 0, $keyUp, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 8
        [ConfirwaNative]::keybd_event($vkAlt, 0, $keyUp, [UIntPtr]::Zero)
        [ConfirwaNative]::keybd_event($vkCtrl, 0, $keyUp, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 20
        $fgAfter = [ConfirwaNative]::GetForegroundWindow()
        if ($fgAfter -ne $fgBefore) {
            $script:lastTerminalHwnd = $fgAfter
        }
    } catch {
        Write-ConfirwaError ("tab switch failed: card={0}, err={1}" -f $TabIndex, ([string]$_.Exception))
    }
}

function Invoke-CardTabSwitch {
    param($Sender)
    if (-not $Sender) {
        return
    }
    $card = Get-CardBySender -Sender $Sender
    if ($card -and (Is-ClickSuppressed -Form $card.Form)) {
        return
    }

    $idx = 0
    try {
        $idx = [int]$Sender.Tag
    } catch {
        $idx = 0
    }
    if ($idx -le 0 -and $Sender -is [System.Windows.Forms.Control]) {
        try {
            $frm = $Sender.FindForm()
            if ($frm) {
                $idx = [int]$frm.Tag
            }
        } catch {
            $idx = 0
        }
    }
    if ($idx -le 0) {
        return
    }
    Send-TerminalTabHotkey -TabIndex $idx
}

function Get-CardBySender {
    param($Sender)
    if (-not $Sender) {
        return $null
    }
    $frm = $null
    if ($Sender -is [System.Windows.Forms.Form]) {
        $frm = $Sender
    } elseif ($Sender -is [System.Windows.Forms.Control]) {
        try {
            $frm = $Sender.FindForm()
        } catch {
            $frm = $null
        }
    }
    if (-not $frm) {
        return $null
    }
    foreach ($c in @($script:cards.Values)) {
        if ($c -and $c.Form -and $c.Form -eq $frm) {
            return $c
        }
    }
    return $null
}

function Cycle-CardHotkeyBinding {
    param($Sender)
    $card = Get-CardBySender -Sender $Sender
    if (-not $card) {
        return
    }

    $title = [string]$card.Title
    if ([string]::IsNullOrWhiteSpace($title)) {
        return
    }
    $key = Normalize-TabTitleKey -Title $title
    $current = 0
    if ($script:titleTabMap.ContainsKey($key)) {
        $current = [int]$script:titleTabMap[$key]
    }
    if ($current -le 0) {
        try { $current = [int]$card.Form.Tag } catch { $current = 1 }
    }
    if ($current -le 0) { $current = 1 }

    $next = (($current % 9) + 1)
    $otherKey = ""
    foreach ($k in @($script:titleTabMap.Keys)) {
        if ($k -eq $key) { continue }
        if ([int]$script:titleTabMap[$k] -eq $next) {
            $otherKey = [string]$k
            break
        }
    }
    $script:titleTabMap[$key] = $next
    if (-not [string]::IsNullOrWhiteSpace($otherKey)) {
        $script:titleTabMap[$otherKey] = $current
    }
    Save-TabMap
}

function Clamp-CardLocation {
    param(
        [System.Drawing.Point] $Point,
        [System.Drawing.Size] $Size,
        [System.Drawing.Rectangle] $Area
    )
    $x = [int]$Point.X
    $y = [int]$Point.Y
    if ($x -lt $Area.Left) { $x = $Area.Left }
    if ($y -lt $Area.Top) { $y = $Area.Top }
    if (($x + $Size.Width) -gt $Area.Right) { $x = [int]($Area.Right - $Size.Width) }
    if (($y + $Size.Height) -gt $Area.Bottom) { $y = [int]($Area.Bottom - $Size.Height) }
    if ($x -lt $Area.Left) { $x = $Area.Left }
    if ($y -lt $Area.Top) { $y = $Area.Top }
    return (New-Object System.Drawing.Point($x, $y))
}

function Start-CardDrag {
    param($Sender, $EventArgs)
    if (-not $Sender -or -not $EventArgs) {
        return
    }
    if ($EventArgs.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $cursor = [System.Windows.Forms.Cursor]::Position
        $script:groupDragState = [PSCustomObject]@{
            Active = $true
            StartCursor = $cursor
            StartOffsetX = [int]$script:anchorOffsetX
            StartOffsetY = [int]$script:anchorOffsetY
            Moved = $false
        }
        return
    }
    if ($EventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) {
        return
    }
    $card = Get-CardBySender -Sender $Sender
    if (-not $card -or -not $card.Form) {
        return
    }
    $form = $card.Form
    $key = [string]$form.Handle.ToInt64()
    $cursor = [System.Windows.Forms.Cursor]::Position
    $script:dragState[$key] = [PSCustomObject]@{
        Active = $true
        StartCursor = $cursor
        StartLocation = $form.Location
        Moved = $false
        CardIndex = [int]$card.Index
        ModelKey = [string]$card.ModelKey
    }
    try { $form.BringToFront() } catch {}
}

function Move-CardDrag {
    param($Sender, $EventArgs)
    if ($script:groupDragState -and $script:groupDragState.Active) {
        $cursor = [System.Windows.Forms.Cursor]::Position
        $dx = [int]($cursor.X - $script:groupDragState.StartCursor.X)
        $dy = [int]($cursor.Y - $script:groupDragState.StartCursor.Y)
        if ([Math]::Abs($dx) + [Math]::Abs($dy) -ge 3) {
            $script:groupDragState.Moved = $true
        }
        if ($script:groupDragState.Moved) {
            $nextX = [int]($script:groupDragState.StartOffsetX + $dx)
            $nextY = [int]($script:groupDragState.StartOffsetY + $dy)
            $script:anchorOffsetX = [Math]::Max(-4000, [Math]::Min(4000, $nextX))
            $script:anchorOffsetY = [Math]::Max(-4000, [Math]::Min(4000, $nextY))
            Update-CardLayout
        }
        return
    }
    if (-not $Sender) {
        return
    }
    $card = Get-CardBySender -Sender $Sender
    if (-not $card -or -not $card.Form) {
        return
    }
    $form = $card.Form
    $key = [string]$form.Handle.ToInt64()
    if (-not $script:dragState.ContainsKey($key)) {
        return
    }
    $state = $script:dragState[$key]
    if (-not $state -or -not $state.Active) {
        return
    }

    $cursor = [System.Windows.Forms.Cursor]::Position
    $dx = [int]($cursor.X - $state.StartCursor.X)
    $dy = [int]($cursor.Y - $state.StartCursor.Y)
    if ([Math]::Abs($dx) + [Math]::Abs($dy) -ge 3) {
        $state.Moved = $true
    }
    if (-not $state.Moved) {
        return
    }

    # Fixed slot layout: drag only changes order on release, never free-position cards.
}

function Get-DropTargetCardIndex {
    param([System.Drawing.Point] $Cursor)
    $bestIdx = 0
    $bestDist = [double]::MaxValue
    foreach ($idx in @($script:cards.Keys | Sort-Object)) {
        $card = $script:cards[[int]$idx]
        if (-not $card -or -not $card.Form) { continue }
        $loc = $card.Form.Location
        $size = $card.Form.Size
        $cx = [double]($loc.X + ($size.Width / 2.0))
        $cy = [double]($loc.Y + ($size.Height / 2.0))
        $dx = $cx - [double]$Cursor.X
        $dy = $cy - [double]$Cursor.Y
        $dist = ($dx * $dx) + ($dy * $dy)
        if ($dist -lt $bestDist) {
            $bestDist = $dist
            $bestIdx = [int]$idx
        }
    }
    return $bestIdx
}

function Move-SlotOrderByKey {
    param(
        [string] $ModelKey,
        [int] $TargetIndex
    )
    $nk = Normalize-ModelKey -Key $ModelKey
    if ([string]::IsNullOrWhiteSpace($nk) -or $nk.StartsWith("__empty_")) {
        return
    }
    if ($TargetIndex -le 0) {
        return
    }

    $visible = @()
    foreach ($idx in @($script:cards.Keys | Sort-Object)) {
        $card = $script:cards[[int]$idx]
        if (-not $card) { continue }
        $k = Normalize-ModelKey -Key ([string]$card.ModelKey)
        if ([string]::IsNullOrWhiteSpace($k) -or $k.StartsWith("__empty_")) { continue }
        if (-not ($visible -contains $k)) {
            $visible += $k
        }
    }
    if (-not ($visible -contains $nk)) {
        $visible += $nk
    }

    $without = @()
    foreach ($k in $visible) {
        if ($k -ne $nk) {
            $without += $k
        }
    }
    $insertAt = [Math]::Max(0, [Math]::Min($TargetIndex - 1, $without.Count))
    $newOrder = @()
    for ($i = 0; $i -lt $without.Count; $i++) {
        if ($i -eq $insertAt) {
            $newOrder += $nk
        }
        $newOrder += $without[$i]
    }
    if ($insertAt -ge $without.Count) {
        $newOrder += $nk
    }
    $script:slotOrder = @($newOrder)
    Save-SlotOrder
}

function End-CardDrag {
    param($Sender, $EventArgs)
    if ($script:groupDragState -and $script:groupDragState.Active) {
        $moved = $false
        try { $moved = [bool]$script:groupDragState.Moved } catch { $moved = $false }
        $script:groupDragState = $null
        if ($moved) {
            Save-AnchorOffset
        }
        Update-CardLayout
        return
    }
    if (-not $Sender) {
        return
    }
    $card = Get-CardBySender -Sender $Sender
    if (-not $card -or -not $card.Form) {
        return
    }
    $form = $card.Form
    $key = [string]$form.Handle.ToInt64()
    if (-not $script:dragState.ContainsKey($key)) {
        return
    }
    $state = $script:dragState[$key]
    $script:dragState.Remove($key) | Out-Null
    if (-not $state -or -not $state.Active) {
        return
    }
    if ($state.Moved) {
        $cursor = [System.Windows.Forms.Cursor]::Position
        $targetIdx = Get-DropTargetCardIndex -Cursor $cursor
        if ($targetIdx -le 0) {
            $targetIdx = [int]$state.CardIndex
        }
        Move-SlotOrderByKey -ModelKey ([string]$state.ModelKey) -TargetIndex $targetIdx
        Set-ClickSuppress -Form $form -Milliseconds 420
    }
    Update-CardLayout
}

function Get-CardAnchorInfo {
    param(
        [System.Drawing.Rectangle] $FallbackArea,
        [int] $LeftPad = 6,
        [int] $TopPad = 44
    )

    $area = $FallbackArea
    $x = [int]($area.Left + $LeftPad)
    $y = [int]($area.Top + $TopPad)

    $target = [IntPtr]$script:lastTerminalHwnd
    if ($target -eq [IntPtr]::Zero -or (Is-ConfirwaHandle -Hwnd $target)) {
        $target = Find-TerminalWindowHandle
    }
    if ($target -ne [IntPtr]::Zero) {
        try {
            $rect = New-Object ConfirwaRect
            if ([ConfirwaNative]::GetWindowRect($target, [ref]$rect)) {
                $winPt = New-Object System.Drawing.Point([int]$rect.Left, [int]$rect.Top)
                $area = [System.Windows.Forms.Screen]::FromPoint($winPt).WorkingArea
                $x = [int]([Math]::Max($area.Left + 2, [int]$rect.Left + $LeftPad))
                $y = [int]([Math]::Max($area.Top + 2, [int]$rect.Top + $TopPad))
            }
        } catch {
        }
    }

    $x = [int]($x + [int]$script:anchorOffsetX)
    $y = [int]($y + [int]$script:anchorOffsetY)

    return [PSCustomObject]@{
        Area = $area
        X = $x
        Y = $y
    }
}

function Get-RoundedGraphicsPath {
    param(
        [System.Drawing.Rectangle] $Rect,
        [int] $Radius = 8
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $r = [Math]::Max(1, $Radius)
    $d = [Math]::Max(2, $r * 2)
    if ($Rect.Width -le ($d + 2) -or $Rect.Height -le ($d + 2)) {
        $path.AddRectangle($Rect)
        return $path
    }

    $arc = New-Object System.Drawing.Rectangle($Rect.X, $Rect.Y, $d, $d)
    $path.AddArc($arc, 180, 90)
    $arc.X = $Rect.Right - $d
    $path.AddArc($arc, 270, 90)
    $arc.Y = $Rect.Bottom - $d
    $path.AddArc($arc, 0, 90)
    $arc.X = $Rect.Left
    $path.AddArc($arc, 90, 90)
    $path.CloseFigure()
    return $path
}

function Set-ControlRoundedRegion {
    param(
        [System.Windows.Forms.Control] $Control,
        [int] $Radius = 8
    )

    if (-not $Control) {
        return
    }
    if ($Control.Width -le 1 -or $Control.Height -le 1) {
        return
    }

    $path = $null
    $newRegion = $null
    $oldRegion = $null
    try {
        $rect = New-Object System.Drawing.Rectangle(0, 0, [int]$Control.Width, [int]$Control.Height)
        $path = Get-RoundedGraphicsPath -Rect $rect -Radius $Radius
        $newRegion = New-Object System.Drawing.Region($path)
        try { $oldRegion = $Control.Region } catch { $oldRegion = $null }
        $Control.Region = $newRegion
        $newRegion = $null
        if ($oldRegion) {
            try { $oldRegion.Dispose() } catch {}
        }
    } catch {
    } finally {
        if ($newRegion) {
            try { $newRegion.Dispose() } catch {}
        }
        if ($path) {
            try { $path.Dispose() } catch {}
        }
    }
}

function Set-CardSpeechBubbleLayout {
    param(
        [object] $Card,
        [string] $SpeechText
    )

    if (-not $Card -or -not $Card.Form -or -not $Card.Bubble -or -not $Card.Picture -or -not $Card.Label) {
        return $false
    }

    $cardWidth = 156
    $cardMinHeight = 126
    $cardMaxHeight = 206
    $bubbleMinHeight = 24
    $bubbleMaxHeight = 112
    $tailHeight = 6
    $footerHeight = 14
    $pictureHeight = 73

    try { if ($Card.PSObject.Properties.Name -contains "CardWidth") { $cardWidth = [int]$Card.CardWidth } } catch {}
    try { if ($Card.PSObject.Properties.Name -contains "CardMinHeight") { $cardMinHeight = [int]$Card.CardMinHeight } } catch {}
    try { if ($Card.PSObject.Properties.Name -contains "CardMaxHeight") { $cardMaxHeight = [int]$Card.CardMaxHeight } } catch {}
    try { if ($Card.PSObject.Properties.Name -contains "BubbleMinHeight") { $bubbleMinHeight = [int]$Card.BubbleMinHeight } } catch {}
    try { if ($Card.PSObject.Properties.Name -contains "BubbleMaxHeight") { $bubbleMaxHeight = [int]$Card.BubbleMaxHeight } } catch {}
    try { if ($Card.PSObject.Properties.Name -contains "TailHeight") { $tailHeight = [int]$Card.TailHeight } } catch {}
    try { if ($Card.PSObject.Properties.Name -contains "FooterHeight") { $footerHeight = [int]$Card.FooterHeight } } catch {}
    try { if ($Card.PSObject.Properties.Name -contains "PictureHeight") { $pictureHeight = [int]$Card.PictureHeight } } catch {}

    if ($cardMaxHeight -lt $cardMinHeight) { $cardMaxHeight = $cardMinHeight }
    if ($bubbleMaxHeight -lt $bubbleMinHeight) { $bubbleMaxHeight = $bubbleMinHeight }
    if ($pictureHeight -lt 24) { $pictureHeight = 24 }

    $text = Normalize-SpeechText -Text $SpeechText
    if ([string]::IsNullOrWhiteSpace($text)) {
        $text = "..."
    }

    $bubbleWidth = [Math]::Max(36, $cardWidth)
    $innerWidth = [Math]::Max(24, $bubbleWidth - 10)
    $desiredBubbleHeight = $bubbleMinHeight
    $g = $null
    $fmt = $null
    try {
        $g = $Card.Bubble.CreateGraphics()
        $fmt = New-Object System.Drawing.StringFormat
        $fmt.Trimming = [System.Drawing.StringTrimming]::Word
        $sizeF = $g.MeasureString($text, $Card.Bubble.Font, $innerWidth, $fmt)
        $desiredBubbleHeight = [int][Math]::Ceiling([double]$sizeF.Height + 8.0)
    } catch {
        $desiredBubbleHeight = $bubbleMinHeight
    } finally {
        if ($fmt) { $fmt.Dispose() }
        if ($g) { $g.Dispose() }
    }
    if ($desiredBubbleHeight -lt $bubbleMinHeight) { $desiredBubbleHeight = $bubbleMinHeight }
    $bubbleHeight = [Math]::Min($bubbleMaxHeight, $desiredBubbleHeight)

    $desiredCardHeight = $bubbleHeight + $tailHeight + $footerHeight + $pictureHeight
    $cardHeight = [Math]::Max($cardMinHeight, $desiredCardHeight)
    $cardHeight = [Math]::Min($cardMaxHeight, $cardHeight)

    $layoutChanged = $false
    if ([int]$Card.Form.Width -ne $cardWidth -or [int]$Card.Form.Height -ne $cardHeight) {
        $Card.Form.Size = New-Object System.Drawing.Size($cardWidth, $cardHeight)
        $layoutChanged = $true
    }

    $bubbleRectChanged = ([int]$Card.Bubble.Left -ne 0 -or [int]$Card.Bubble.Top -ne 0 -or [int]$Card.Bubble.Width -ne $bubbleWidth -or [int]$Card.Bubble.Height -ne $bubbleHeight)
    if ($bubbleRectChanged) {
        $Card.Bubble.SetBounds(0, 0, $bubbleWidth, $bubbleHeight)
        $layoutChanged = $true
    }
    if ($bubbleRectChanged -or $null -eq $Card.Bubble.Region) {
        Set-ControlRoundedRegion -Control $Card.Bubble -Radius 8
    }

    if ($Card.PSObject.Properties.Name -contains "BubbleTail" -and $Card.BubbleTail) {
        $tailRectChanged = ([int]$Card.BubbleTail.Left -ne 0 -or [int]$Card.BubbleTail.Top -ne $bubbleHeight -or [int]$Card.BubbleTail.Width -ne $cardWidth -or [int]$Card.BubbleTail.Height -ne $tailHeight)
        if ($tailRectChanged) {
            $Card.BubbleTail.SetBounds(0, $bubbleHeight, $cardWidth, $tailHeight)
            $layoutChanged = $true
        }
    }

    $pictureTop = $bubbleHeight + $tailHeight
    $pictureRectChanged = ([int]$Card.Picture.Left -ne 0 -or [int]$Card.Picture.Top -ne $pictureTop -or [int]$Card.Picture.Width -ne $cardWidth -or [int]$Card.Picture.Height -ne $pictureHeight)
    if ($pictureRectChanged) {
        $Card.Picture.SetBounds(0, $pictureTop, $cardWidth, $pictureHeight)
        $layoutChanged = $true
    }

    $labelTop = $bubbleHeight + $tailHeight + $pictureHeight
    $labelRectChanged = ([int]$Card.Label.Left -ne 0 -or [int]$Card.Label.Top -ne $labelTop -or [int]$Card.Label.Width -ne $cardWidth -or [int]$Card.Label.Height -ne $footerHeight)
    if ($labelRectChanged) {
        $Card.Label.SetBounds(0, $labelTop, $cardWidth, $footerHeight)
        $layoutChanged = $true
    }

    $currentBubbleText = ""
    try { $currentBubbleText = [string]$Card.Bubble.AccessibleDescription } catch { $currentBubbleText = "" }
    if ($currentBubbleText -ne $text) {
        $Card.Bubble.AccessibleDescription = $text
        $Card.Bubble.Text = ""
        try { $Card.Bubble.Invalidate() } catch {}
    }
    return $layoutChanged
}

function New-Card {
    param([int] $Index)

    $cardWidth = 156
    $cardMinHeight = 126
    $cardMaxHeight = 206
    $bubbleMinHeight = 24
    $bubbleMaxHeight = 112
    $tailHeight = 6
    $footerHeight = 14
    $pictureHeight = [Math]::Max(24, $cardMinHeight - $bubbleMinHeight - $tailHeight - $footerHeight)
    $transparentKey = [System.Drawing.Color]::FromArgb(1, 1, 1)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "codex - confirwa"
    $form.Width = $cardWidth
    $form.Height = $cardMinHeight
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $form.ControlBox = $false
    $form.BackColor = $transparentKey
    $form.TransparencyKey = $transparentKey

    $picture = New-Object System.Windows.Forms.PictureBox
    $picture.SetBounds(0, $bubbleMinHeight + $tailHeight, $cardWidth, $pictureHeight)
    $picture.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $picture.BackColor = $transparentKey

    $bubble = New-Object System.Windows.Forms.Label
    $bubble.SetBounds(0, 0, $cardWidth, $bubbleMinHeight)
    $bubble.ForeColor = [System.Drawing.Color]::Transparent
    $bubble.BackColor = [System.Drawing.Color]::Transparent
    $bubble.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $bubble.AutoSize = $false
    $bubble.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
    $bubble.Padding = New-Object System.Windows.Forms.Padding(5, 2, 5, 1)
    $bubble.UseCompatibleTextRendering = $false
    try {
        $bubble.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8.5, [System.Drawing.FontStyle]::Bold)
    } catch {
        $bubble.Font = New-Object System.Drawing.Font("Segoe UI", 8.3, [System.Drawing.FontStyle]::Bold)
    }
    $bubble.Text = ""
    $bubble.AccessibleDescription = "..."

    $bubbleTail = New-Object System.Windows.Forms.Label
    $bubbleTail.SetBounds(0, $bubbleMinHeight, $cardWidth, $tailHeight)
    $bubbleTail.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
    $bubbleTail.ForeColor = [System.Drawing.Color]::Transparent
    $bubbleTail.BackColor = [System.Drawing.Color]::Transparent
    $bubbleTail.Text = ""

    $label = New-Object System.Windows.Forms.Label
    $label.SetBounds(0, $bubbleMinHeight + $tailHeight + $pictureHeight, $cardWidth, $footerHeight)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $label.ForeColor = [System.Drawing.Color]::Transparent
    $label.BackColor = [System.Drawing.Color]::Transparent
    $label.UseCompatibleTextRendering = $false
    $label.AutoEllipsis = $false
    $label.Text = ""
    $label.AccessibleDescription = "codex: idle"
    try {
        $label.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8.2, [System.Drawing.FontStyle]::Regular)
    } catch {
        try {
            $label.Font = New-Object System.Drawing.Font("Segoe UI", 8.2, [System.Drawing.FontStyle]::Regular)
        } catch {
        }
    }

    $mouseHandler = {
        param($sender, $e)
        if (-not $e) {
            return
        }
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            Invoke-CardTabSwitch -Sender $sender
            return
        }
    }
    $dragDownHandler = {
        param($sender, $e)
        Start-CardDrag -Sender $sender -EventArgs $e
    }
    $dragMoveHandler = {
        param($sender, $e)
        Move-CardDrag -Sender $sender -EventArgs $e
    }
    $dragUpHandler = {
        param($sender, $e)
        End-CardDrag -Sender $sender -EventArgs $e
    }
    $bubblePaintHandler = {
        param($sender, $e)
        if (-not $sender -or -not $e) {
            return
        }
        $path = $null
        $fillBrush = $null
        $pen = $null
        $fmt = $null
        $textBrush = $null
        try {
            $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
            $rect = New-Object System.Drawing.Rectangle(1, 1, [Math]::Max(2, $sender.Width - 3), [Math]::Max(2, $sender.Height - 3))
            $path = Get-RoundedGraphicsPath -Rect $rect -Radius 8
            $fillBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 253, 247))
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(46, 48, 54), 2.0)
            $e.Graphics.FillPath($fillBrush, $path)
            $e.Graphics.DrawPath($pen, $path)

            $txt = ""
            try { $txt = [string]$sender.AccessibleDescription } catch { $txt = "" }
            if (-not [string]::IsNullOrWhiteSpace($txt)) {
                $fmt = New-Object System.Drawing.StringFormat
                $fmt.Alignment = [System.Drawing.StringAlignment]::Near
                $fmt.LineAlignment = [System.Drawing.StringAlignment]::Near
                $fmt.Trimming = [System.Drawing.StringTrimming]::EllipsisWord
                $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(24, 26, 30))
                $textRect = New-Object System.Drawing.RectangleF(6.0, 2.0, [Math]::Max(1.0, ($sender.Width - 12.0)), [Math]::Max(1.0, ($sender.Height - 4.0)))
                $e.Graphics.DrawString($txt, $sender.Font, $textBrush, $textRect, $fmt)
            }
        } catch {
        } finally {
            if ($textBrush) { $textBrush.Dispose() }
            if ($fmt) { $fmt.Dispose() }
            if ($fillBrush) { $fillBrush.Dispose() }
            if ($pen) { $pen.Dispose() }
            if ($path) { $path.Dispose() }
        }
    }
    $bubbleTailPaintHandler = {
        param($sender, $e)
        if (-not $sender -or -not $e) {
            return
        }
        $fillBrush = $null
        $pen = $null
        try {
            $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $w = [int]$sender.Width
            $h = [int]$sender.Height
            if ($w -le 4 -or $h -le 1) {
                return
            }
            $cx = [int][Math]::Floor($w / 2)
            $half = [int][Math]::Max(3, [Math]::Floor($h * 0.95))
            $baseY = 0
            $tipY = [int]([Math]::Max(1, $h - 1))
            $points = [System.Drawing.Point[]]@(
                (New-Object System.Drawing.Point(($cx - $half), $baseY)),
                (New-Object System.Drawing.Point(($cx + $half), $baseY)),
                (New-Object System.Drawing.Point($cx, $tipY))
            )
            $fillBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 253, 247))
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(46, 48, 54), 1.0)
            $e.Graphics.FillPolygon($fillBrush, $points)
            $e.Graphics.DrawPolygon($pen, $points)
        } catch {
        } finally {
            if ($pen) { $pen.Dispose() }
            if ($fillBrush) { $fillBrush.Dispose() }
        }
    }
    $labelPaintHandler = {
        param($sender, $e)
        if (-not $sender -or -not $e) {
            return
        }
        $txt = ""
        try {
            $txt = [string]$sender.AccessibleDescription
        } catch {
            $txt = ""
        }
        if ([string]::IsNullOrWhiteSpace($txt)) {
            try { $txt = [string]$sender.Text } catch { $txt = "" }
        }
        if ([string]::IsNullOrWhiteSpace($txt)) {
            return
        }

        $fmt = $null
        $brush = $null
        try {
            $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::None
            $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::SingleBitPerPixelGridFit
            $fmt = New-Object System.Drawing.StringFormat
            $fmt.Alignment = [System.Drawing.StringAlignment]::Center
            $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
            $fmt.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
            $fmt.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(194, 202, 214))
            $r = $sender.ClientRectangle
            $textRect = New-Object System.Drawing.RectangleF($r.X, $r.Y, $r.Width, $r.Height)
            $e.Graphics.DrawString($txt, $sender.Font, $brush, $textRect, $fmt)
        } finally {
            if ($brush) { $brush.Dispose() }
            if ($fmt) { $fmt.Dispose() }
        }
    }
    $form.Tag = [int]$Index
    $picture.Tag = [int]$Index
    $bubble.Tag = [int]$Index
    $bubbleTail.Tag = [int]$Index
    $label.Tag = [int]$Index
    $form.Add_MouseClick($mouseHandler)
    $picture.Add_MouseClick($mouseHandler)
    $bubble.Add_MouseClick($mouseHandler)
    $bubbleTail.Add_MouseClick($mouseHandler)
    $label.Add_MouseClick($mouseHandler)
    $form.Add_MouseDown($dragDownHandler)
    $picture.Add_MouseDown($dragDownHandler)
    $bubble.Add_MouseDown($dragDownHandler)
    $bubbleTail.Add_MouseDown($dragDownHandler)
    $label.Add_MouseDown($dragDownHandler)
    $form.Add_MouseMove($dragMoveHandler)
    $picture.Add_MouseMove($dragMoveHandler)
    $bubble.Add_MouseMove($dragMoveHandler)
    $bubbleTail.Add_MouseMove($dragMoveHandler)
    $label.Add_MouseMove($dragMoveHandler)
    $form.Add_MouseUp($dragUpHandler)
    $picture.Add_MouseUp($dragUpHandler)
    $bubble.Add_MouseUp($dragUpHandler)
    $bubbleTail.Add_MouseUp($dragUpHandler)
    $label.Add_MouseUp($dragUpHandler)
    $bubble.Add_Paint($bubblePaintHandler)
    $bubbleTail.Add_Paint($bubbleTailPaintHandler)
    $label.Add_Paint($labelPaintHandler)

    $form.Controls.Add($picture)
    $form.Controls.Add($bubbleTail)
    $form.Controls.Add($bubble)
    $form.Controls.Add($label)

    $script:cards[$Index] = [PSCustomObject]@{
        Index = [int]$Index
        Form = $form
        Picture = $picture
        Bubble = $bubble
        BubbleTail = $bubbleTail
        Label = $label
        ImagePath = ""
        Title = "codex"
        State = "idle"
        Speech = "..."
        CardWidth = [int]$cardWidth
        CardMinHeight = [int]$cardMinHeight
        CardMaxHeight = [int]$cardMaxHeight
        BubbleMinHeight = [int]$bubbleMinHeight
        BubbleMaxHeight = [int]$bubbleMaxHeight
        TailHeight = [int]$tailHeight
        FooterHeight = [int]$footerHeight
        PictureHeight = [int]$pictureHeight
        ModelKey = ("__empty_{0}" -f $Index)
    }
    [void](Set-CardSpeechBubbleLayout -Card $script:cards[$Index] -SpeechText "...")
    Set-ControlRoundedRegion -Control $bubble -Radius 8
    [void]$form.Show()
}

function Remove-Card {
    param([int] $Index)
    if (-not $script:cards.ContainsKey($Index)) {
        return
    }
    $card = $script:cards[$Index]
    try {
        $card.Form.Close()
        $card.Form.Dispose()
    } catch {
    }
    $script:cards.Remove($Index) | Out-Null
}

function Update-CardLayout {
    $keys = @($script:cards.Keys | Sort-Object)
    if ($keys.Count -eq 0) {
        return
    }

    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $sample = $script:cards[[int]$keys[0]]
    $cardWidth = [int]$sample.Form.Width
    $cardHeight = [int]$sample.Form.Height
    $gap = 6
    $anchor = Get-CardAnchorInfo -FallbackArea $screen -LeftPad $gap -TopPad 44
    $targetArea = [System.Drawing.Rectangle]$anchor.Area
    $startX = [int]$anchor.X
    $startY = [int]$anchor.Y
    $maxPerRow = [Math]::Max(1, [Math]::Floor(($targetArea.Width + $gap) / ($cardWidth + $gap)))
    $perRow = [Math]::Min($keys.Count, $maxPerRow)
    if ($perRow -le 0) {
        $perRow = 1
    }

    $rowHeights = @{}
    $maxRow = 0
    for ($i = 0; $i -lt $keys.Count; $i++) {
        $idx = [int]$keys[$i]
        $row = [int][Math]::Floor($i / $perRow)
        if ($row -gt $maxRow) {
            $maxRow = $row
        }
        $h = $cardHeight
        try {
            $h = [int]$script:cards[$idx].Form.Height
        } catch {
            $h = $cardHeight
        }
        if (-not $rowHeights.ContainsKey($row) -or [int]$rowHeights[$row] -lt $h) {
            $rowHeights[$row] = $h
        }
    }

    $rowTop = @{}
    $cursorY = $startY
    for ($r = 0; $r -le $maxRow; $r++) {
        $rowTop[$r] = [int]$cursorY
        $rh = $cardHeight
        if ($rowHeights.ContainsKey($r)) {
            $rh = [int]$rowHeights[$r]
        }
        $cursorY += ($rh + $gap)
    }

    $groupWidth = [int](($perRow * $cardWidth) + (($perRow - 1) * $gap))
    if ($groupWidth -lt $cardWidth) {
        $groupWidth = $cardWidth
    }
    $groupHeight = 0
    for ($r = 0; $r -le $maxRow; $r++) {
        $rh = $cardHeight
        if ($rowHeights.ContainsKey($r)) {
            $rh = [int]$rowHeights[$r]
        }
        $groupHeight += $rh
        if ($r -lt $maxRow) {
            $groupHeight += $gap
        }
    }

    $minX = [int]$targetArea.Left
    $maxX = [int]($targetArea.Right - $groupWidth)
    if ($maxX -lt $minX) {
        $maxX = $minX
    }
    $startX = [int][Math]::Max($minX, [Math]::Min($maxX, $startX))

    $minY = [int]$targetArea.Top
    $maxY = [int]($targetArea.Bottom - $groupHeight)
    if ($maxY -lt $minY) {
        $maxY = $minY
    }
    $startY = [int][Math]::Max($minY, [Math]::Min($maxY, $startY))

    $rowTop = @{}
    $cursorY = $startY
    for ($r = 0; $r -le $maxRow; $r++) {
        $rowTop[$r] = [int]$cursorY
        $rh = $cardHeight
        if ($rowHeights.ContainsKey($r)) {
            $rh = [int]$rowHeights[$r]
        }
        $cursorY += ($rh + $gap)
    }

    for ($i = 0; $i -lt $keys.Count; $i++) {
        $idx = [int]$keys[$i]
        $row = [int][Math]::Floor($i / $perRow)
        $col = [int]($i % $perRow)
        $rowHeight = $cardHeight
        if ($rowHeights.ContainsKey($row)) {
            $rowHeight = [int]$rowHeights[$row]
        }
        $itemHeight = $cardHeight
        try {
            $itemHeight = [int]$script:cards[$idx].Form.Height
        } catch {
            $itemHeight = $cardHeight
        }
        $pt = New-Object System.Drawing.Point(
            [int]($startX + ($col * ($cardWidth + $gap))),
            [int]($rowTop[$row] + ($rowHeight - $itemHeight))
        )
        $clamped = Clamp-CardLocation -Point $pt -Size $script:cards[$idx].Form.Size -Area $targetArea
        $script:cards[$idx].Form.Location = $clamped
    }
}

function Set-CardCount {
    param([int] $Target)
    if ($Target -lt 0) {
        $Target = 0
    }

    $current = $script:cards.Count
    while ($current -lt $Target) {
        $next = 1
        if ($script:cards.Count -gt 0) {
            $next = ([int](@($script:cards.Keys | Measure-Object -Maximum).Maximum)) + 1
        }
        New-Card -Index $next
        $current = $script:cards.Count
    }

    while ($current -gt $Target) {
        $keys = @($script:cards.Keys | Sort-Object -Descending)
        Remove-Card -Index ([int]$keys[0])
        $current = $script:cards.Count
    }

    Update-CardLayout
}

function Apply-CardVisuals {
    param([object[]] $Models)

    $keys = @($script:cards.Keys | Sort-Object)
    $needsRelayout = $false
    for ($i = 0; $i -lt $keys.Count; $i++) {
        $idx = [int]$keys[$i]
        $card = $script:cards[$idx]
        $model = if ($i -lt $Models.Count) { $Models[$i] } else { [PSCustomObject]@{ Title = "codex"; State = "idle"; Speech = "..."; HotkeyIndex = $idx; Key = ("__empty_{0}" -f $idx) } }

        $title = [string]$model.Title
        if ([string]::IsNullOrWhiteSpace($title)) { $title = "codex" }
        $state = [string]$model.State
        if ($state -ne "running" -and $state -ne "approval" -and $state -ne "reconnecting" -and $state -ne "silent" -and $state -ne "idle") {
            $state = "idle"
        }

        $imgPath = Get-ImagePath $state
        $needImageUpdate = $card.ImagePath -ne $imgPath
        if (-not $needImageUpdate -and $null -eq $card.Picture.ImageLocation -and (Test-Path $imgPath)) {
            $needImageUpdate = $true
        }

        if ($needImageUpdate) {
            if (Test-Path $imgPath) {
                try {
                    $card.Picture.ImageLocation = $imgPath
                } catch {
                    $card.Picture.ImageLocation = $null
                }
            } else {
                $card.Picture.ImageLocation = $null
            }
            $card.ImagePath = $imgPath
        }

        if ($card.Title -ne $title) {
            $card.Title = $title
            $card.Form.Text = "{0} - confirwa" -f $title
        }

        $modelKey = [string]$model.Key
        if ([string]::IsNullOrWhiteSpace($modelKey)) {
            $modelKey = ("__empty_{0}" -f $idx)
        }
        $card.ModelKey = $modelKey

        $hotkeyIndex = 0
        try {
            $hotkeyIndex = [int]$model.HotkeyIndex
        } catch {
            $hotkeyIndex = 0
        }
        if ($hotkeyIndex -le 0) {
            $hotkeyIndex = $idx
        }
        $card.Form.Tag = $hotkeyIndex
        $card.Picture.Tag = $hotkeyIndex
        $card.Bubble.Tag = $hotkeyIndex
        if ($card.PSObject.Properties.Name -contains "BubbleTail" -and $card.BubbleTail) {
            $card.BubbleTail.Tag = $hotkeyIndex
        }
        $card.Label.Tag = $hotkeyIndex

        $speech = Normalize-SpeechText -Text ([string]$model.Speech)
        if ([string]::IsNullOrWhiteSpace($speech)) {
            $speech = "..."
        }
        $layoutChanged = Set-CardSpeechBubbleLayout -Card $card -SpeechText $speech
        if ($layoutChanged) {
            $needsRelayout = $true
        }
        if ($card.Speech -ne $speech) {
            $card.Speech = $speech
        }

        $labelText = "[{0}] {1}: {2}" -f $hotkeyIndex, $title, (Get-StateText $state)
        $currentLabelText = ""
        try { $currentLabelText = [string]$card.Label.AccessibleDescription } catch { $currentLabelText = "" }
        if ($currentLabelText -ne $labelText) {
            $card.Label.AccessibleDescription = $labelText
            $card.Label.Text = ""
            try { $card.Label.Invalidate() } catch {}
        }
        $card.State = $state
    }
    if ($needsRelayout) {
        Update-CardLayout
    }
}

function Update-GlobalStateFromModels {
    param([object[]] $Models)

    $state = "idle"
    foreach ($m in $Models) {
        if ($m.State -eq "approval") {
            $state = "approval"
            break
        }
        if ($m.State -eq "reconnecting") {
            $state = "reconnecting"
            continue
        }
        if ($m.State -eq "silent") {
            if ($state -eq "idle" -or $state -eq "running") {
                $state = "silent"
            }
            continue
        }
        if ($m.State -eq "running") {
            if ($state -eq "idle") {
                $state = "running"
            }
        }
    }
    Set-CurrentState $state
}

$driver = New-Object System.Windows.Forms.Form
$driver.Text = "confirwa-driver"
$driver.ShowInTaskbar = $false
$driver.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$driver.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$driver.Location = New-Object System.Drawing.Point(-32000, -32000)
$driver.Size = New-Object System.Drawing.Size(1, 1)
$driver.Opacity = 0
$script:driverForm = $driver

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [Math]::Max(200, $PollMs)
$timer.Add_Tick({
    try {
        Refresh-LastTerminalWindowFromForeground
        Poll-Sessions

        $agentCount = Get-CodexAgentCount
        Set-CardCount -Target $agentCount

        if ($script:cards.Count -gt 0) {
            $models = Get-CardModels -Count $script:cards.Count
            Apply-CardVisuals -Models $models
            Update-GlobalStateFromModels -Models $models
        } else {
            Set-CurrentState "idle"
        }
    } catch {
        Write-ConfirwaError ([string]$_.Exception)
    }
})

$driver.Add_Shown({
    $driver.Hide()
    $timer.Start()
})

$driver.Add_FormClosing({
    $timer.Stop()
    foreach ($idx in @($script:cards.Keys | Sort-Object)) {
        Remove-Card -Index ([int]$idx)
    }
})

[void]$driver.Show()
[System.Windows.Forms.Application]::Run($driver)
