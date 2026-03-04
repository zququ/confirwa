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
            ApprovalPending = $false
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
    }
    if ($Reconnecting) {
        if ($when -gt $script:sessionStatus[$Path].LastReconnectUtc) {
            $script:sessionStatus[$Path].LastReconnectUtc = $when
        }
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
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Line)) {
        $s = $Line.ToLowerInvariant()
        if ($s -match '"requires_approval"\s*:\s*true') { return $true }
        if ($s -match '"require_escalated"') { return $true }
        if ($s -match '"approval_request"') { return $true }
        if ($s -match 'yes,\s*proceed\s*\(y\)') { return $true }
        if ($s -match '^\s*[›>]*\s*1\.\s*yes,\s*proceed\s*\(y\)') { return $true }
        if ($s -match '^\s*[2２]\.\s*yes,\s*and\s*don''t ask again') { return $true }
        if ($s -match '\(\s*y\s*\)' -and $s -match 'yes' -and $s -match 'proceed') { return $true }
        if ($s -match 'do you want me to proceed') { return $true }
        if ($s -match 'want me to proceed') { return $true }
        if ($s -match 'can i proceed') { return $true }
        if ($s -match 'may i proceed') { return $true }
        if ($s -match 'should i proceed') { return $true }
        if ($s -match 'proceed\?') { return $true }
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
            if ($pType -match "reconnect|reconnecting|connection_lost|connection_restored") {
                return $true
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Line)) {
        $s = $Line.ToLowerInvariant()
        if ($s -match '"reconnecting"') { return $true }
        if ($s -match '"reconnect"') { return $true }
        if ($s -match 'connection_lost') { return $true }
        if ($s -match 'connection_restored') { return $true }
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
            Clear-SessionApprovalPending -Path $Path
            break
        }
        "task_complete" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            Clear-SessionApprovalPending -Path $Path
            break
        }
        "turn_aborted" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            Clear-SessionApprovalPending -Path $Path
            break
        }
        "task_failed" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            Clear-SessionApprovalPending -Path $Path
            break
        }
        "task_cancelled" {
            if ($turns.ContainsKey($turnId)) {
                $turns.Remove($turnId) | Out-Null
            }
            Clear-SessionApprovalPending -Path $Path
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
    $isEventMsg = ($Line -match '"type"\s*:\s*"event_msg"')
    $isResponseItem = ($Line -match '"type"\s*:\s*"response_item"')
    $isFunctionCall = ($Line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"function_call"')
    $isFunctionCallOutput = ($Line -match '"payload"\s*:\s*\{\s*"type"\s*:\s*"(function_call_output|custom_tool_call_output)"')
    $hasApprovalHint = ($Line -match '"requires_approval"\s*:\s*true' -or
                        $Line -match '"require_escalated"' -or
                        $Line -match '"approval_request"' -or
                        $Line -match 'yes,\s*proceed\s*\(y\)' -or
                        $Line -match '^\s*[›>]*\s*1\.\s*yes,\s*proceed\s*\(y\)' -or
                        $Line -match '\(\s*y\s*\)' -or
                        $Line -match 'do you want me to proceed' -or
                        $Line -match 'want me to proceed' -or
                        $Line -match 'can i proceed' -or
                        $Line -match 'may i proceed' -or
                        $Line -match 'should i proceed' -or
                        $Line -match 'proceed\?')

    if ($isFunctionCallOutput) {
        Clear-SessionApprovalPending -Path $Path
    }

    if (-not $isEventMsg -and -not $isResponseItem) {
        if (Is-ReconnectingSignal -Obj $null -Line $Line) {
            Mark-SessionEvent -Path $Path -Reconnecting -AtUtc $eventUtc
            return
        }
        if (Is-ApprovalSignal -Obj $null -Line $Line) {
            Mark-SessionEvent -Path $Path -Approval -AtUtc $eventUtc
            return
        }
        return
    }

    if ($isEventMsg) {
        Update-SessionTurnStateFromLine -Path $Path -Line $Line -AtUtc $eventUtc

        if (Is-ApprovalSignal -Obj $null -Line $Line) {
            Mark-SessionEvent -Path $Path -Approval -AtUtc $eventUtc
            return
        }

        if (Is-ReconnectingSignal -Obj $null -Line $Line) {
            Mark-SessionEvent -Path $Path -Reconnecting -AtUtc $eventUtc
            return
        }

        Mark-SessionEvent -Path $Path -AtUtc $eventUtc
        return
    }

    if ($isFunctionCall -or $hasApprovalHint) {
        if (Is-ApprovalSignal -Obj $null -Line $Line) {
            Mark-SessionEvent -Path $Path -Approval -AtUtc $eventUtc
            return
        }
    }

    if (Is-ReconnectingSignal -Obj $null -Line $Line) {
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

            if (Is-ApprovalSignal -Obj $null -Line $line) {
                Mark-SessionEvent -Path $Path -Approval -AtUtc $eventUtc
            }
            if (Is-ReconnectingSignal -Obj $null -Line $line) {
                Mark-SessionEvent -Path $Path -Reconnecting -AtUtc $eventUtc
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
    if ($script:recentSessionInfos.Count -eq 0) {
        return
    }

    $remaining = 260
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

    $processCount = 0
    try {
        $processCount = @(Get-Process -Name codex -ErrorAction SilentlyContinue).Count
    } catch {
        $processCount = 0
    }

    $freshSessionCount = 0
    try {
        Refresh-SessionSnapshot
        if ($script:recentSessionInfos.Count -gt 0) {
            $freshSessionCount = @(
                $script:recentSessionInfos |
                    Where-Object { (([DateTime]::UtcNow - ([DateTime]$_.LastWriteTimeUtc).ToUniversalTime()).TotalSeconds -le 3600.0) } |
                    Select-Object -First 12
            ).Count
        }
    } catch {
        $freshSessionCount = 0
    }

    $count = $processCount
    if ($count -le 0) {
        $count = $freshSessionCount
    }
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
    $reconnectHoldSec = [Math]::Max(30.0, $IdleSeconds * 18.0)
    $runningWindowSec = [Math]::Max(30.0, $IdleSeconds * 20.0)
    $status = $null
    if ($script:sessionStatus.ContainsKey($Path)) {
        $status = $script:sessionStatus[$Path]
    }

    $hasActiveTurn = $false
    $activeStartUtc = [DateTime]::MinValue
    $lastWriteUtc = ([DateTime]$LastWriteTimeUtc).ToUniversalTime()
    $writeAgeSec = ($now - $lastWriteUtc).TotalSeconds
    $activeTurnMaxSilenceSec = [Math]::Max(7200.0, $IdleSeconds * 4800.0)
    if ($status -and $null -ne $status.ActiveTurns -and $status.ActiveTurns.Count -gt 0) {
        if ($writeAgeSec -le $activeTurnMaxSilenceSec) {
            $hasActiveTurn = $true
            $activeStartUtc = [DateTime]::UtcNow
            foreach ($k in @($status.ActiveTurns.Keys)) {
                $dt = [DateTime]$status.ActiveTurns[$k]
                if ($dt -lt $activeStartUtc) {
                    $activeStartUtc = $dt
                }
            }
        } else {
            $status.ActiveTurns = @{}
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
        }
        $status.ApprovalPending = $false
    }

    if ($status -and $status.LastApprovalUtc -ne [DateTime]::MinValue) {
        $approvalAge = ($now - $status.LastApprovalUtc).TotalSeconds
        if ($approvalAge -le $approvalHoldSec) {
            return [PSCustomObject]@{
                State = "approval"
                AgeSec = [int][Math]::Floor([Math]::Max(0.0, $approvalAge))
            }
        }
    }

    if ($status -and $status.LastReconnectUtc -ne [DateTime]::MinValue) {
        $reconnectAge = ($now - $status.LastReconnectUtc).TotalSeconds
        if ($reconnectAge -le $reconnectHoldSec) {
            return [PSCustomObject]@{
                State = "reconnecting"
                AgeSec = [int][Math]::Floor([Math]::Max(0.0, $reconnectAge))
            }
        }
    }

    if ($hasActiveTurn) {
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
    $nowUtc = [DateTime]::UtcNow

    $visibleWindowSec = 21600.0
    $candidateInfos = @(
        $script:recentSessionInfos |
            Where-Object {
                $ageSec = ($nowUtc - ([DateTime]$_.LastWriteTimeUtc).ToUniversalTime()).TotalSeconds
                if ($ageSec -le $visibleWindowSec) {
                    return $true
                }
                $p = [string]$_.Path
                if ($script:sessionStatus.ContainsKey($p) -and $null -ne $script:sessionStatus[$p].ActiveTurns -and $script:sessionStatus[$p].ActiveTurns.Count -gt 0) {
                    return $true
                }
                return $false
            }
    )
    if ($candidateInfos.Count -eq 0) {
        $candidateInfos = @($script:recentSessionInfos | Select-Object -First 12)
    }

    $orderedInfos = @(
        $candidateInfos |
            Sort-Object `
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

    $picked = @()
    $usedPath = @{}
    $seenTitle = @{}

    # first pass: prefer different opened directories so 3S/cal are separated
    foreach ($info in $orderedInfos) {
        $title = [string]$info.Title
        if ([string]::IsNullOrWhiteSpace($title)) { $title = "codex" }
        $titleKey = $title.ToLowerInvariant()
        if ($seenTitle.ContainsKey($titleKey)) {
            continue
        }
        $picked += $info
        $usedPath[[string]$info.Path] = $true
        $seenTitle[$titleKey] = $true
        if ($picked.Count -ge $Count) {
            break
        }
    }

    # second pass: fill remaining slots with next most recent sessions
    if ($picked.Count -lt $Count) {
        foreach ($info in $orderedInfos) {
            $path = [string]$info.Path
            if ($usedPath.ContainsKey($path)) {
                continue
            }
            $picked += $info
            $usedPath[$path] = $true
            if ($picked.Count -ge $Count) {
                break
            }
        }
    }

    $selected = @($picked | Select-Object -First $Count)
    $rawModels = @()
    foreach ($info in $selected) {
        $title = [string]$info.Title
        if ([string]::IsNullOrWhiteSpace($title)) {
            $title = "codex"
        }
        $stateInfo = Get-SessionStateInfo -Path ([string]$info.Path) -LastWriteTimeUtc ([DateTime]$info.LastWriteTimeUtc) -FileLength ([int64]$info.FileLength)
        $rawModels += [PSCustomObject]@{
            Title = $title
            State = [string]$stateInfo.State
            AgeSec = [int]$stateInfo.AgeSec
        }
    }

    while ($rawModels.Count -lt $Count) {
        $rawModels += [PSCustomObject]@{
            Title = "codex"
            State = "idle"
            AgeSec = 0
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
        }
    }
    return @($result)
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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:cards = @{}

function New-Card {
    param([int] $Index)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "codex - confirwa"
    $form.Width = 128
    $form.Height = 84
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $form.ControlBox = $false

    $picture = New-Object System.Windows.Forms.PictureBox
    $picture.Dock = [System.Windows.Forms.DockStyle]::Fill
    $picture.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $picture.BackColor = [System.Drawing.Color]::Black

    $label = New-Object System.Windows.Forms.Label
    $label.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $label.Height = 12
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $label.ForeColor = [System.Drawing.Color]::White
    $label.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $label.Text = "codex: idle"

    $form.Controls.Add($picture)
    $form.Controls.Add($label)

    $script:cards[$Index] = [PSCustomObject]@{
        Form = $form
        Picture = $picture
        Label = $label
        ImagePath = ""
        Title = "codex"
        State = "idle"
    }
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
    $cardWidth = 128
    $cardHeight = 84
    $gap = 6
    $perRow = [Math]::Max(1, [Math]::Floor(($screen.Width - $gap) / ($cardWidth + $gap)))

    for ($i = 0; $i -lt $keys.Count; $i++) {
        $idx = [int]$keys[$i]
        $row = [Math]::Floor($i / $perRow)
        $col = $i % $perRow
        $x = [int]($screen.Right - (($col + 1) * ($cardWidth + $gap)))
        $y = [int]($screen.Bottom - (($row + 1) * ($cardHeight + $gap)))
        if ($x -lt $screen.Left) { $x = $screen.Left }
        if ($y -lt $screen.Top) { $y = $screen.Top }
        $script:cards[$idx].Form.Location = New-Object System.Drawing.Point($x, $y)
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
    for ($i = 0; $i -lt $keys.Count; $i++) {
        $idx = [int]$keys[$i]
        $card = $script:cards[$idx]
        $model = if ($i -lt $Models.Count) { $Models[$i] } else { [PSCustomObject]@{ Title = "codex"; State = "idle" } }

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

        $labelText = "{0}: {1}" -f $title, (Get-StateText $state)
        if ($card.Label.Text -ne $labelText) {
            $card.Label.Text = $labelText
        }
        $card.State = $state
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

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [Math]::Max(200, $PollMs)
$timer.Add_Tick({
    try {
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
