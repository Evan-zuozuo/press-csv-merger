#Requires -Version 5.1
# ============================================================
#  Press CSV Merger  V1.0
#
#  Merges the many small CSV files exported by the press every
#  minute into one file per [date + machine + type], and can
#  safely delete the merged source files.
#
#  Environment: built-in PowerShell 5.1 on Windows 10 / 11
#               no Excel, no network, nothing to install
#
#  Encoding: pure ASCII (safe on Japanese / Chinese / English Windows)
#
#  Safety design:
#    S1  Merge phase is strictly read-only on source files
#    S2  Safe time window: only files older than N days are touched
#    S3  Write to .tmp first; promote only after row-count check
#    S4  Per-group atomic: merge -> verify -> delete that group only
#    S5  Never deletes unless -DeleteSource / clean mode + confirm
#    S6  Preview mode: statistics only, no write, no delete
#    S7  Single file read failure -> log and skip, run continues
#    S8  Full run log + deleted-file list (plain text, no disk cost)
#    S9  No non-CSV touching, no registry, no system change, no network
#    S10 Whitelist check before every delete (name pattern + inside source)
#    S11 Idempotent: re-runs only append new files, never overwrite
#
#  System protection (top priority: never disturb the press):
#    P1  Process priority lowered to BelowNormal
#    P2  Disk/CPU throttling (sleep every N files)
#    P3  Single instance lock
#    P4  Global time limit; when it stops, nothing is deleted
#    P5  Emergency stop switch: STOP.txt
#    P6  Scheduled task runs hidden (no window on site)
# ============================================================

param(
    [string]$Src,
    [string]$Out,
    [string]$Mode,
    [int]$SafeDays,
    [int]$MinAgeMinutes,
    [int]$MaxFiles,
    [int]$TimeoutMinutes,
    [int]$ThrottleMs,
    [switch]$DeleteSource,
    [switch]$Yes,
    [switch]$NoPause
)

# ======================== Constants and defaults ========================

$script:Version   = '1.0'
$script:LogLines  = New-Object System.Collections.Generic.List[string]
$script:Deleted   = New-Object System.Collections.Generic.List[string]
$script:FailList  = New-Object System.Collections.Generic.List[string]

# File / folder names (all ASCII on purpose)
$script:SettingsName = '0_Settings_EDIT_ME.txt'
$script:LogFolder    = 'Logs'
$script:MergedFolder = '_Merged'
$script:StopFile     = 'STOP.txt'
$script:LockFile     = '_running.lock'
$script:AutoBat      = '4_AutoRun.bat'

if ([string]::IsNullOrWhiteSpace($Mode)) { $Mode = 'Merge' }
if (-not $PSBoundParameters.ContainsKey('SafeDays'))       { $SafeDays = 1 }
if (-not $PSBoundParameters.ContainsKey('MinAgeMinutes'))  { $MinAgeMinutes = 10 }
if (-not $PSBoundParameters.ContainsKey('MaxFiles'))       { $MaxFiles = 0 }
if (-not $PSBoundParameters.ContainsKey('TimeoutMinutes')) { $TimeoutMinutes = 30 }
if (-not $PSBoundParameters.ContainsKey('ThrottleMs'))     { $ThrottleMs = 30 }
$Mode = $Mode.Trim().ToLower()

# Snapshot for the failure report
$script:RunEnv = @{
    Mode = $Mode; Src = ''; Out = ''; SafeDays = $SafeDays
    MinAgeMinutes = $MinAgeMinutes; MaxFiles = $MaxFiles
    TimeoutMinutes = $TimeoutMinutes; ThrottleMs = $ThrottleMs
    Delete = $false; Rows = 0; Deleted = 0
}
$script:LockPath  = ''
$script:Deadline  = $null
$script:Throttle  = 0
$script:Stopped   = $false

# ======================== Basic helpers ========================

function Write-Line {
    param([string]$Text = '', [string]$Color = 'White')
    if ($Color -eq 'White') { Write-Host $Text }
    else { Write-Host $Text -ForegroundColor $Color }
}

function Add-Log {
    param([string]$Text, [string]$Level = 'INFO')
    $stamp = Get-Date -Format 'HH:mm:ss'
    $script:LogLines.Add("[$stamp][$Level] $Text")
    switch ($Level) {
        'WARN'  { Write-Host "  ! $Text" -ForegroundColor Yellow }
        'ERROR' { Write-Host "  X $Text" -ForegroundColor Red }
        default { Write-Host "  . $Text" }
    }
}

function Get-ScriptDir {
    if ($PSScriptRoot) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return Split-Path -Parent $MyInvocation.MyCommand.Path }
    return (Get-Location).Path
}

function Read-TextLinesAuto {
    param([string]$Path)
    # Encoding auto-detect: UTF-8(BOM) -> system ANSI -> UTF-8(no BOM)
    $raw = [System.IO.File]::ReadAllBytes($Path)
    if ($raw.Length -ge 3 -and $raw[0] -eq 0xEF -and $raw[1] -eq 0xBB -and $raw[2] -eq 0xBF) {
        return [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    }
    $lines = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::Default)
    if ((($lines -join '') -match [char]0xFFFD)) {
        return [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    }
    return $lines
}

function Read-ConfigFile {
    param([string]$Path)
    $cfg = @{}
    if (-not (Test-Path $Path)) { return $cfg }
    try {
        $lines = Read-TextLinesAuto -Path $Path
        foreach ($line in $lines) {
            if ($line -match '^\s*([^=;#]+?)\s*=\s*(.*)$') {
                $cfg[$Matches[1].Trim()] = $Matches[2].Trim()
            }
        }
    } catch {
        Add-Log -Text "Cannot read settings file: $Path - $($_.Exception.Message)" -Level WARN
    }
    return $cfg
}

function Save-ConfigFile {
    param([string]$Path, [string]$SrcPath, [string]$OutPath, [int]$Days)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('# Press CSV Merger - settings')
    [void]$sb.AppendLine('# Edit with Notepad and save. ASCII path recommended.')
    [void]$sb.AppendLine("Src=$SrcPath")
    [void]$sb.AppendLine("Out=$OutPath")
    [void]$sb.AppendLine("SafeDays=$Days")
    try {
        $enc = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($Path, $sb.ToString(), $enc)
    } catch {
        Add-Log -Text "Cannot write settings file: $($_.Exception.Message)" -Level WARN
    }
}

function Get-FileBody {
    param([string]$Path)
    # Returns @{ Header=..; Row=.. }, or $null on failure
    try {
        $text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    } catch {
        try { $text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default) }
        catch { return $null }
    }
    $header = $null
    $row    = $null
    foreach ($ln in ($text -split "`r?`n")) {
        $t = $ln.Trim()
        if ($t -eq '') { continue }
        if ($null -eq $header) { $header = $t.TrimEnd(',') ; continue }
        $row = $ln.TrimEnd("`r").TrimEnd(',')
        break
    }
    if ($null -eq $header -or $null -eq $row) { return $null }
    return @{ Header = $header; Row = $row }
}

function Split-CsvLine {
    param([string]$Line)
    # Simple split (data is numeric/alphanumeric, no quoted fields)
    return @($Line -split ',')
}

function Invoke-DelChunk {
    param([string[]]$Paths)
    # cmd "del" batch delete: measured ~25x faster than one-by-one File.Delete
    $argList = '/c del /f /q'
    foreach ($p in $Paths) { $argList += ' "' + $p + '"' }
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'cmd.exe'
        $psi.Arguments = $argList
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.WindowStyle = 'Hidden'
        $proc = [System.Diagnostics.Process]::Start($psi)
        if ($null -ne $proc) { $proc.WaitForExit(120000) | Out-Null }
        return $true
    } catch { return $false }
}

function Invoke-FastDelete {
    param(
        [System.Collections.Generic.List[string]]$Paths,
        [string]$StopDir = ''
    )
    $chunk  = New-Object System.Collections.Generic.List[string]
    $argLen = 0
    $round  = 0
    foreach ($p in $Paths) {
        # Emergency stop: check every 25 batches (~1500 files)
        if ($StopDir -ne '' -and $round -gt 0 -and $round % 25 -eq 0 -and (Test-StopRequested -Dir $StopDir)) {
            Write-Host '    STOP.txt detected, delete halted'
            return
        }
        $chunk.Add($p)
        $argLen += $p.Length + 4
        if ($chunk.Count -ge 60 -or $argLen -gt 6000) {
            $ok = Invoke-DelChunk -Paths $chunk.ToArray()
            if (-not $ok) {
                foreach ($q in $chunk) { try { [System.IO.File]::Delete($q) } catch {} }
            }
            $round++
            if ($round % 25 -eq 0) { Write-Host "    Deleting $([Math]::Min($round * 60, $Paths.Count)) / $($Paths.Count)" }
            # Throttle: rest between batches, do not saturate the disk
            if ($script:Throttle -gt 0) { Start-Sleep -Milliseconds $script:Throttle }
            $chunk.Clear(); $argLen = 0
        }
    }
    if ($chunk.Count -gt 0) {
        $ok = Invoke-DelChunk -Paths $chunk.ToArray()
        if (-not $ok) {
            foreach ($q in $chunk) { try { [System.IO.File]::Delete($q) } catch {} }
        }
    }
}

function Get-RemainingSet {
    param([string]$Src, [DateTime]$Cutoff)
    # Re-scan the source folder, return paths of files still present
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    try {
        $files = [System.IO.Directory]::GetFiles($Src, '*', [System.IO.SearchOption]::AllDirectories)
    } catch { return ,$set }
    foreach ($p in $files) {
        $ext = [System.IO.Path]::GetExtension($p)
        if ($ext -eq '' -or -not ($ext -ieq '.csv')) { continue }
        $nm = [System.IO.Path]::GetFileName($p)
        if ($nm -notmatch '^(\d{14})_(.+?)_(Utility|WORKINFO)\.CSV$') { continue }
        $fd = [DateTime]::MinValue
        if (-not [DateTime]::TryParseExact($Matches[1].Substring(0, 8), 'yyyyMMdd', $null,
                [System.Globalization.DateTimeStyles]::None, [ref]$fd)) { continue }
        if ($fd -ge $Cutoff) { continue }
        [void]$set.Add($p)
    }
    return ,$set
}

function Set-LowPriority {
    # Drop to BelowNormal: yields CPU to the press program when busy
    try {
        $proc = [System.Diagnostics.Process]::GetCurrentProcess()
        $proc.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::BelowNormal
        return $true
    } catch { return $false }
}

function Acquire-RunLock {
    param([string]$Dir)
    # Single instance lock: no two runs touching the same files
    $lock = Join-Path $Dir $script:LockFile
    if (Test-Path -LiteralPath $lock) {
        try {
            $age = (Get-Date) - (Get-Item -LiteralPath $lock).LastWriteTime
            if ($age.TotalHours -lt 4) { return $null }
            Remove-Item -LiteralPath $lock -Force -ErrorAction SilentlyContinue
        } catch { return $null }
    }
    try {
        $fs = [System.IO.File]::Open($lock, [System.IO.FileMode]::CreateNew,
            [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $w = New-Object System.IO.StreamWriter($fs)
        $w.WriteLine("PID=$PID")
        $w.WriteLine("Start=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $w.Flush(); $w.Close(); $fs.Close()
        return $lock
    } catch { return $null }
}

function Release-RunLock {
    param([string]$LockPath)
    if ([string]::IsNullOrWhiteSpace($LockPath)) { return }
    try { if (Test-Path -LiteralPath $LockPath) { Remove-Item -LiteralPath $LockPath -Force -ErrorAction Stop } } catch {}
}

function Test-StopRequested {
    param([string]$Dir)
    # Emergency stop: engineer drops a STOP.txt in the folder
    try { return (Test-Path -LiteralPath (Join-Path $Dir $script:StopFile)) } catch { return $false }
}

function Write-FailLog {
    param(
        [string]$ScriptDir,
        [string]$Message,
        [string]$Stack = '',
        [int]$ExitCode = 0
    )
    # Failure report for the engineer: must say WHERE it broke
    $dir = Join-Path $ScriptDir $script:LogFolder
    if (-not (Test-Path $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null } catch { return }
    }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $fp = Join-Path $dir ("FAILURE_$stamp.txt")

    $L = New-Object System.Collections.Generic.List[string]
    $L.Add('============================================')
    $L.Add('  Press CSV Merger - FAILURE REPORT')
    $L.Add('============================================')
    $L.Add('')
    $L.Add("Time        : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $L.Add("Computer    : $env:COMPUTERNAME")
    $L.Add("User        : $env:USERDOMAIN\$env:USERNAME")
    $L.Add("PowerShell  : $($PSVersionTable.PSVersion.ToString())")
    $L.Add('')
    $L.Add('--- Run settings ---')
    $L.Add("Mode        : $($script:RunEnv.Mode)")
    $L.Add("Source      : $($script:RunEnv.Src)")
    $L.Add("Output      : $($script:RunEnv.Out)")
    $L.Add("SafeDays    : $($script:RunEnv.SafeDays)")
    $L.Add("MinAgeMin   : $($script:RunEnv.MinAgeMinutes)")
    $L.Add("MaxFiles    : $($script:RunEnv.MaxFiles)")
    $L.Add("TimeoutMin  : $($script:RunEnv.TimeoutMinutes)")
    $L.Add("Delete mode : $($script:RunEnv.Delete)")
    $L.Add('')
    $L.Add('--- Progress before stop ---')
    $L.Add("Rows merged : $($script:RunEnv.Rows)")
    $L.Add("Files deleted: $($script:RunEnv.Deleted)")
    $L.Add('')
    $L.Add('--- Error ---')
    $L.Add($Message)
    if ($Stack -ne '') { $L.Add(''); $L.Add('Location:'); $L.Add($Stack) }
    if ($ExitCode -ne 0) { $L.Add("Exit code   : $ExitCode") }
    $L.Add('')
    $L.Add('============================================')
    $L.Add('  What to do (for the engineer)')
    $L.Add('============================================')
    $L.Add('1. Send this whole file to the engineer.')
    $L.Add('2. Source files are safe: this tool merges first and')
    $L.Add('   deletes only after the row-count check passed.')
    $L.Add('3. Just run it again; it resumes, no dup, no loss.')
    $L.Add('4. If the same error repeats, usual causes:')
    $L.Add('   - Source path missing or renamed -> fix Src= in settings')
    $L.Add('   - No write permission on output -> change Out=')
    $L.Add('   - Source folder locked by another program -> retry later')
    $L.Add('   - Code.ps1 or a .bat was deleted -> re-copy from backup')
    $L.Add('')

    try {
        [System.IO.File]::WriteAllLines($fp, $L, (New-Object System.Text.UTF8Encoding($true)))
        Write-Host "  Failure report written: $fp" -ForegroundColor Yellow
    } catch {}
}

function Write-LogFile {
    param([string]$ScriptDir)
    $dir = Join-Path $ScriptDir $script:LogFolder
    if (-not (Test-Path $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null } catch { return }
    }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $fp = Join-Path $dir ("RUNLOG_$stamp.txt")
    $i = 1
    while (Test-Path $fp) { $i++; $fp = Join-Path $dir ("RUNLOG_$stamp" + "_$i.txt") }
    try {
        [System.IO.File]::WriteAllLines($fp, $script:LogLines, (New-Object System.Text.UTF8Encoding($true)))
        Write-Line "  Log file: $fp" Gray
    } catch {
        Write-Line "  ! Cannot write log: $($_.Exception.Message)" Yellow
    }
    if ($script:Deleted.Count -gt 0) {
        $fp2 = Join-Path $dir ("DELETED_$stamp.txt")
        $j = 1
        while (Test-Path $fp2) { $j++; $fp2 = Join-Path $dir ("DELETED_$stamp" + "_$j.txt") }
        try {
            [System.IO.File]::WriteAllLines($fp2, $script:Deleted, (New-Object System.Text.UTF8Encoding($true)))
            Write-Line "  Deleted list: $fp2" Gray
        } catch {}
    }
}

# ======================== Main ========================

function Main {

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $scriptDir = Get-ScriptDir
    Set-Location $scriptDir

    # ---------- Safety: low priority + single instance lock ----------
    $null = Set-LowPriority

    $script:LockPath = Acquire-RunLock -Dir $scriptDir
    if ([string]::IsNullOrWhiteSpace($script:LockPath)) {
        Write-Line ''
        Write-Line '  Another instance is already running. Exiting.' Yellow
        Write-Line '  (Protection: never let two runs touch the same files)' Gray
        Write-Line ''
        return
    }

    if (Test-StopRequested -Dir $scriptDir) {
        Write-Line ''
        Write-Line "  $($script:StopFile) detected. Nothing will run." Yellow
        Write-Line '  (Emergency stop switch. Delete that file to resume.)' Gray
        Write-Line ''
        Release-RunLock -LockPath $script:LockPath
        $script:LockPath = ''
        return
    }

    $script:Deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $script:Throttle = $ThrottleMs

    Write-Line ''
    Write-Line '============================================' Cyan
    Write-Line "   Press CSV Merger  V$($script:Version)" Cyan
    Write-Line '============================================' Cyan
    Write-Line ''

    # ---------- Mode ----------
    $isPreview = $false
    $doDelete  = $false
    switch ($Mode) {
        'preview'       { $isPreview = $true }
        'pe'            { $isPreview = $true }
        'merge'         { }
        'mergeandclean' { $doDelete = $true }
        'clean'         { $doDelete = $true }
        default {
            Write-Line "  X Unknown mode: $Mode (use: Preview / Merge / MergeAndClean)" Red
            if (-not $NoPause) { Read-Host '  Press Enter to close' }
            return
        }
    }
    if ($DeleteSource) { $doDelete = $true }

    # ---------- Settings ----------
    $cfgPath = Join-Path $scriptDir $script:SettingsName
    $cfg = Read-ConfigFile -Path $cfgPath

    if ([string]::IsNullOrWhiteSpace($Src)) { $Src = $cfg['Src'] }
    if ([string]::IsNullOrWhiteSpace($Out)) { $Out = $cfg['Out'] }
    if (-not $PSBoundParameters.ContainsKey('SafeDays') -and $cfg.ContainsKey('SafeDays')) {
        $v = 0
        if ([int]::TryParse($cfg['SafeDays'], [ref]$v)) { $SafeDays = $v }
    }

    if ([string]::IsNullOrWhiteSpace($Src)) {
        if ($NoPause) {
            # Auto mode: never wait for input, or the schedule hangs forever
            Write-Line '  X Auto mode: source folder not configured' Red
            Write-Line "     Edit $($script:SettingsName) and set Src= first" Yellow
            Add-Log -Text 'Auto mode: Src not set in settings, exiting' -Level ERROR
            Write-LogFile -ScriptDir $scriptDir
            return
        }
        Write-Line '  [First use] Set the source folder' Yellow
        Write-Line '  (the folder where the press exports CSV, e.g. D:\PressData\Export)' Gray
        $in = Read-Host '  Source folder path'
        $Src = $in.Trim().Trim('"')
    }

    if ([string]::IsNullOrWhiteSpace($Src)) {
        Write-Line '  X No source folder. Exiting.' Red
        if (-not $NoPause) { Read-Host '  Press Enter to close' }
        return
    }

    if (-not (Test-Path $Src)) {
        Write-Line "  X Source folder not found: $Src" Red
        if (-not $NoPause) { Read-Host '  Press Enter to close' }
        return
    }
    $Src = (Resolve-Path $Src).Path

    if ([string]::IsNullOrWhiteSpace($Out)) {
        $Out = Join-Path $Src $script:MergedFolder
    } elseif (-not [System.IO.Path]::IsPathRooted($Out)) {
        $Out = Join-Path $scriptDir $Out
    }
    try { $outFull = [System.IO.Path]::GetFullPath($Out).TrimEnd('\') } catch { $outFull = '' }

    $script:RunEnv.Src = $Src
    $script:RunEnv.Out = $Out
    $script:RunEnv.Delete = $doDelete

    # ---------- Environment ----------
    Write-Line '  [Environment]' Yellow
    Write-Line "  V PowerShell $($PSVersionTable.PSVersion.ToString())" Green
    Write-Line "  V Source folder: $Src" Green

    $outInsideSrc = $false
    if ($outFull -ne '') {
        $srcFull = $Src.TrimEnd('\')
        if ($outFull -ieq $srcFull -or $outFull.StartsWith($srcFull + '\', [StringComparison]::OrdinalIgnoreCase)) {
            $outInsideSrc = $true
        }
    }

    $cutoff = (Get-Date).Date.AddDays(-($SafeDays - 1))
    $modeText = if ($isPreview) { 'PREVIEW (statistics only, no write, no delete)' }
                elseif ($doDelete) { 'MERGE + CLEAN (delete sources after merge)' }
                else { 'MERGE (keep sources)' }

    Write-Line "  Mode: $modeText" White
    Write-Line "  Safe window: only files before $($cutoff.ToString('yyyy-MM-dd')) 00:00 (SafeDays=$SafeDays)" White
    if ($MinAgeMinutes -gt 0) {
        Write-Line "  File cooling: files newer than $MinAgeMinutes minutes are skipped" White
    }
    if ($ThrottleMs -gt 0) {
        Write-Line "  Throttle: sleep $ThrottleMs ms every 200 files (low system load)" White
    }
    if ($outInsideSrc) {
        Write-Line '  Output folder is inside source: it will be excluded from scanning' Gray
    }
    Write-Line ''

    # ---------- Scan ----------
    Write-Line '  [Scanning CSV]' Yellow
    Add-Log -Text "===== START mode=$Mode src=$Src ====="

    $swScan = [System.Diagnostics.Stopwatch]::StartNew()
    $allFiles = @()
    try {
        $allFiles = [System.IO.Directory]::GetFiles($Src, '*', [System.IO.SearchOption]::AllDirectories)
    } catch {
        Write-Line "  X Scan failed: $($_.Exception.Message)" Red
        if (-not $NoPause) { Read-Host '  Press Enter to close' }
        return
    }
    $swScan.Stop()

    $groups      = @{}
    $nameDup     = @{}
    $nTotal      = 0
    $nCsv        = 0
    $nSkipToday  = 0
    $nUnknown    = 0
    $nOutScope   = 0
    $nSkipFresh  = 0
    $nSkipOwn    = 0

    foreach ($p in $allFiles) {
        $nTotal++

        # Exclude our own output folder (never re-merge our results)
        if ($outFull -ne '') {
            $pd = [System.IO.Path]::GetDirectoryName($p)
            if ($pd) {
                $pd = $pd.TrimEnd('\')
                if ($pd -ieq $outFull -or $pd.StartsWith($outFull + '\', [StringComparison]::OrdinalIgnoreCase)) {
                    $nSkipOwn++
                    continue
                }
            }
        }

        $ext = [System.IO.Path]::GetExtension($p)
        if ($ext -eq '') { continue }
        if (-not ($ext -ieq '.csv')) { continue }
        $nCsv++

        $name = [System.IO.Path]::GetFileName($p)
        if ($name -notmatch '^(\d{14})_(.+?)_(Utility|WORKINFO)\.CSV$') {
            $nUnknown++
            continue
        }
        $ts   = $Matches[1]
        $mach = $Matches[2].ToUpper()
        $typ  = $Matches[3]

        $fd = [DateTime]::MinValue
        if (-not [DateTime]::TryParseExact($ts.Substring(0, 8), 'yyyyMMdd', $null,
                [System.Globalization.DateTimeStyles]::None, [ref]$fd)) {
            $nUnknown++
            continue
        }
        if ($fd -ge $cutoff) { $nSkipToday++; continue }

        # Cooling: use the timestamp INSIDE THE FILE NAME (= real export time).
        # Never use LastWriteTime: copying/moving a file resets it, which would
        # make old files look brand new and never get processed.
        if ($MinAgeMinutes -gt 0) {
            $fdt = [DateTime]::MinValue
            if ([DateTime]::TryParseExact($ts, 'yyyyMMddHHmmss', $null,
                    [System.Globalization.DateTimeStyles]::None, [ref]$fdt)) {
                if (((Get-Date) - $fdt).TotalMinutes -lt $MinAgeMinutes) { $nSkipFresh++; continue }
            }
        }

        $key = $ts.Substring(0, 8) + '_' + $mach + '_' + $typ

        # Same name in different folders: prefix with parent folder name
        $tag = $name
        if ($nameDup.ContainsKey($name)) {
            $parent = Split-Path -Parent $p
            $pname  = [System.IO.Path]::GetFileName($parent)
            $tag    = $pname + '_' + $name
            Add-Log -Text "Duplicate name in another folder, tag changed to $tag" -Level WARN
        }
        $nameDup[$name] = $true

        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = New-Object System.Collections.Generic.List[object]
        }
        $groups[$key].Add(@{ Path = $p; Tag = $tag; Ts = $ts })
        $nOutScope++
    }

    Write-Line "  Files in folder: $nTotal" White
    if ($nSkipOwn -gt 0)   { Write-Line "  In output folder (excluded): $nSkipOwn" Gray }
    Write-Line "  CSV files: $nCsv" White
    Write-Line "  Name pattern matched: $($nOutScope + $nSkipToday + $nSkipFresh)" White
    if ($nUnknown -gt 0) {
        Write-Line "  ! Name pattern not matched (skipped): $nUnknown" Yellow
        Add-Log -Text "Skipped, bad name pattern: $nUnknown" -Level WARN
    }
    if ($nSkipToday -gt 0) {
        Write-Line "  ! Inside safe window / too new (skipped): $nSkipToday" Yellow
        Add-Log -Text "Skipped, safe window: $nSkipToday" -Level WARN
    }
    if ($nSkipFresh -gt 0) {
        Write-Line "  ! Inside cooling time (skipped): $nSkipFresh" Yellow
        Add-Log -Text "Skipped, cooling: $nSkipFresh" -Level WARN
    }

    # ---------- Batch limit (avoid heavy disk load) ----------
    $nDeferred = 0
    if ($MaxFiles -gt 0 -and $nOutScope -gt $MaxFiles) {
        $kept = @{}
        $cnt = 0
        foreach ($k in ($groups.Keys | Sort-Object)) {
            if ($cnt + $groups[$k].Count -gt $MaxFiles) { break }
            $kept[$k] = $groups[$k]
            $cnt += $groups[$k].Count
        }
        $nDeferred = $groups.Count - $kept.Count
        $groups = $kept
        $nOutScope = $cnt
        Write-Line "  ! Limited to $MaxFiles this run, $nDeferred groups left for next time" Yellow
        Add-Log -Text "Batch limit: $MaxFiles this run, $nDeferred groups deferred" -Level WARN
    }

    Write-Line "  To merge: $nOutScope files in $($groups.Count) groups (scan $($swScan.Elapsed.TotalSeconds.ToString('N1')) s)" Green
    Write-Line ''

    if ($groups.Count -eq 0) {
        Write-Line '  Nothing to merge.' Yellow
        Write-Line '  (If the press export is off and you want today too, use -SafeDays 0)' Gray
        Add-Log -Text 'Nothing to merge'
        Write-LogFile -ScriptDir $scriptDir
        if (-not $NoPause) { Read-Host '  Press Enter to close' }
        return
    }

    # ---------- Preview ----------
    if ($isPreview) {
        Write-Line '  [Preview] These files would be created:' Yellow
        $sample = 0
        foreach ($k in ($groups.Keys | Sort-Object)) {
            $sample++
            if ($sample -le 12) {
                Write-Line ("    {0,-34} {1,7} rows" -f ($k + '.csv'), $groups[$k].Count)
            }
        }
        if ($groups.Count -gt 12) { Write-Line "    ... $($groups.Count - 12) more groups" Gray }
        Write-Line ''
        Write-Line "  Total: $($groups.Count) merged files covering $nOutScope source files" Green
        $estSec = [int][Math]::Round($nOutScope * 0.0025)
        $estText = if ($estSec -ge 90) { ([Math]::Round($estSec / 60, 1).ToString()) + ' min' } else { $estSec.ToString() + ' s' }
        Write-Line "  Estimated merge time: about $estText" Gray
        Add-Log -Text "Preview done: $($groups.Count) groups / $nOutScope files"
        Write-LogFile -ScriptDir $scriptDir
        if (-not $NoPause) { Read-Host '  Press Enter to close' }
        return
    }

    # ---------- Output folder ----------
    if (-not (Test-Path $Out)) {
        try { New-Item -ItemType Directory -Path $Out -Force | Out-Null }
        catch {
            Write-Line "  X Cannot create output folder: $Out" Red
            if (-not $NoPause) { Read-Host '  Press Enter to close' }
            return
        }
    }
    Write-Line "  Output folder: $Out" Green
    Write-Line ''

    # ---------- Merge ----------
    Write-Line '  [Merging]' Yellow
    $swMerge = [System.Diagnostics.Stopwatch]::StartNew()

    $utf8          = New-Object System.Text.UTF8Encoding($false)
    $nGroupOk      = 0
    $nGroupSkip    = 0
    $nNewRows      = 0
    $nReadFail     = 0
    $toDelete      = New-Object System.Collections.Generic.List[string]
    $processed     = 0
    $throttleCount = 0

    foreach ($k in ($groups.Keys | Sort-Object)) {

        # --- Safety gates: time limit / emergency stop ---
        if ($null -ne $script:Deadline -and (Get-Date) -gt $script:Deadline) {
            Write-Line "  ! Time limit ($TimeoutMinutes min) reached, stopping safely" Yellow
            Add-Log -Text "Time limit $TimeoutMinutes min reached, safe stop (rest next run)" -Level WARN
            $script:Stopped = $true
            break
        }
        if (Test-StopRequested -Dir $scriptDir) {
            Write-Line "  ! $($script:StopFile) detected, stopping safely" Yellow
            Add-Log -Text 'STOP.txt detected, safe stop' -Level WARN
            $script:Stopped = $true
            break
        }

        $processed++
        $items = $groups[$k]
        $dst   = Join-Path $Out ($k + '.csv')
        $tmp   = Join-Path $Out ($k + '.csv.tmp')

        # --- Existing merged file: read back what was already merged ---
        $seen = New-Object 'System.Collections.Generic.HashSet[string]'
        $oldDataCount = 0
        $existingHeader = $null
        if (Test-Path $dst) {
            try {
                $rd = New-Object System.IO.StreamReader($dst, $utf8, $true)
                $isFirst = $true
                while ($null -ne ($ln = $rd.ReadLine())) {
                    if ($ln.Trim() -eq '') { continue }
                    if ($isFirst) { $existingHeader = $ln; $isFirst = $false; continue }
                    $oldDataCount++
                    $idx = $ln.LastIndexOf(',')
                    if ($idx -ge 0) { [void]$seen.Add($ln.Substring($idx + 1).Trim()) }
                }
                $rd.Close()
            } catch {
                Add-Log -Text "Cannot read existing merged file, will rebuild: $k - $($_.Exception.Message)" -Level WARN
                $seen = New-Object 'System.Collections.Generic.HashSet[string]'
                $oldDataCount = 0
                $existingHeader = $null
            }
        }

        # --- Filter out files already merged ---
        $newItems = New-Object System.Collections.Generic.List[object]
        foreach ($it in $items) {
            if (-not $seen.Contains($it.Tag)) { $newItems.Add($it) }
        }

        if ($newItems.Count -eq 0) {
            # Already merged before. Data is safely inside the merged file,
            # so in clean mode still queue them for cleanup, otherwise
            # "merged but delete failed" files would stay forever.
            if ($doDelete) {
                foreach ($it in $items) {
                    if ($seen.Contains($it.Tag)) { $toDelete.Add($it.Path) }
                }
            }
            $nGroupSkip++
            continue
        }

        # --- Write .tmp (copy existing content, then append new rows) ---
        $headerToWrite = $null
        $added = 0
        try {
            $wr = New-Object System.IO.StreamWriter($tmp, $false, $utf8)

            if ($null -ne $existingHeader) {
                $wr.WriteLine($existingHeader)
                if (Test-Path $dst) {
                    $rd2 = New-Object System.IO.StreamReader($dst, $utf8, $true)
                    $first = $true
                    while ($null -ne ($ln2 = $rd2.ReadLine())) {
                        if ($ln2.Trim() -eq '') { continue }
                        if ($first) { $first = $false; continue }
                        $wr.WriteLine($ln2)
                    }
                    $rd2.Close()
                }
            }

            foreach ($it in $newItems) {
                # Disk/CPU throttle: rest every 200 files so the press
                # program and operators never feel the load
                if ($script:Throttle -gt 0) {
                    $throttleCount++
                    if ($throttleCount % 200 -eq 0) { Start-Sleep -Milliseconds $script:Throttle }
                }
                # Emergency stop (checked every 500 files to limit disk reads)
                if ($throttleCount % 500 -eq 0 -and (Test-StopRequested -Dir $scriptDir)) {
                    $script:Stopped = $true
                    break
                }

                $body = Get-FileBody -Path $it.Path
                if ($null -eq $body) { $nReadFail++; $script:FailList.Add($it.Path); continue }
                if ($null -eq $headerToWrite) { $headerToWrite = $body.Header }

                if ($null -eq $existingHeader -and $added -eq 0) {
                    $wr.WriteLine($body.Header + ',SrcFile')
                }

                $ncol = ($headerToWrite -split ',').Count
                $vals = Split-CsvLine -Line $body.Row
                if ($vals.Count -gt $ncol) { $vals = $vals[0..($ncol - 1)] }
                while ($vals.Count -lt $ncol) { $vals += '' }
                $wr.WriteLine(($vals -join ',') + ',' + $it.Tag)
                $added++
            }
            $wr.Flush()
            $wr.Close()
        } catch {
            Add-Log -Text "Write failed, group skipped: $k - $($_.Exception.Message)" -Level ERROR
            try { $wr.Close() } catch {}
            if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
            continue
        }

        # Stopped mid-read: this group does not count, drop tmp, delete nothing
        if ($script:Stopped) {
            try { if (Test-Path $tmp) { [System.IO.File]::Delete($tmp) } } catch {}
            Add-Log -Text "STOP: group $k not promoted, source files kept" -Level WARN
            break
        }

        if ($added -eq 0) {
            if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
            continue
        }

        # --- S3 verify: tmp rows must equal existing + added ---
        $tmpData = 0
        try {
            $rd3 = New-Object System.IO.StreamReader($tmp, $utf8, $true)
            $first = $true
            while ($null -ne ($ln3 = $rd3.ReadLine())) {
                if ($ln3.Trim() -eq '') { continue }
                if ($first) { $first = $false; continue }
                $tmpData++
            }
            $rd3.Close()
        } catch {
            Add-Log -Text "Verify read failed, group not promoted: $k" -Level ERROR
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
            continue
        }

        $expect = $oldDataCount + $added
        if ($tmpData -ne $expect) {
            Add-Log -Text "Verify failed, group not promoted: $k (expected $expect rows, got $tmpData). Sources kept." -Level ERROR
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
            continue
        }

        # --- Verified: promote ---
        # Use overwrite, not "delete then move": if the old file is open in
        # Excel or blocked by policy, delete-then-move loses the whole group.
        $ok = $false
        try {
            [System.IO.File]::Copy($tmp, $dst, $true)
            $ok = $true
        } catch {
            try {
                if (Test-Path $dst) { [System.IO.File]::Delete($dst) }
                [System.IO.File]::Move($tmp, $dst)
                $ok = $true
            } catch {
                Add-Log -Text "Promote failed: $k - $($_.Exception.Message). Sources kept." -Level ERROR
            }
        }
        if (-not $ok) { continue }
        try { if (Test-Path $tmp) { [System.IO.File]::Delete($tmp) } } catch {}

        $nGroupOk++
        $nNewRows += $added

        # --- S4 per-group atomic: only queue this group's sources ---
        foreach ($it in $newItems) { $toDelete.Add($it.Path) }

        if ($processed % 20 -eq 0) {
            Write-Host ("    Progress {0}/{1} groups" -f $processed, $groups.Count)
        }
    }

    $swMerge.Stop()
    $script:RunEnv.Rows = $nNewRows
    Write-Line ''
    Write-Line "  Merge done: $nGroupOk groups ok, $nNewRows rows added, $nGroupSkip groups already done" Green
    if ($nReadFail -gt 0) {
        Write-Line "  ! Read failures skipped: $nReadFail (see log; those were NOT deleted)" Yellow
        Add-Log -Text "Read failures: $nReadFail" -Level WARN
    }
    Write-Line "  Merge time: $($swMerge.Elapsed.TotalSeconds.ToString('N1')) s" White

    # ---------- Delete ----------
    $nDeleted = 0
    $nDelFail = 0

    # If stopped by a safety mechanism, do not delete anything this run
    if ($script:Stopped -and $doDelete) {
        Write-Line ''
        Write-Line '  Stopped by safety mechanism: nothing deleted this run. Sources kept.' Yellow
        Add-Log -Text 'Stopped by safety mechanism: no source file deleted this run' -Level WARN
        $doDelete = $false
    }

    if ($doDelete -and $toDelete.Count -gt 0) {

        Write-Line ''
        Write-Line '  [Cleaning source files]' Yellow

        # S10 whitelist + second age check before delete
        $safeToDelete = New-Object System.Collections.Generic.List[string]
        $rootFull = (Resolve-Path $Src).Path.TrimEnd('\')
        foreach ($p in $toDelete) {
            $full = ''
            try { $full = (Resolve-Path $p).Path } catch { $nDelFail++; continue }
            if (-not $full.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) {
                Add-Log -Text "Refuse delete (outside source folder): $full" -Level ERROR
                $nDelFail++; continue
            }
            $nm = [System.IO.Path]::GetFileName($full)
            if ($nm -notmatch '^\d{14}_.+_(Utility|WORKINFO)\.CSV$') {
                Add-Log -Text "Refuse delete (bad name pattern): $nm" -Level ERROR
                $nDelFail++; continue
            }
            # Last line of defense: re-check the timestamp inside the file name.
            # Again not LastWriteTime (copying a file resets it).
            if ($MinAgeMinutes -gt 0 -and $nm -match '^(\d{14})_') {
                $fdt2 = [DateTime]::MinValue
                if ([DateTime]::TryParseExact($Matches[1], 'yyyyMMddHHmmss', $null,
                        [System.Globalization.DateTimeStyles]::None, [ref]$fdt2)) {
                    if (((Get-Date) - $fdt2).TotalMinutes -lt $MinAgeMinutes) {
                        if ($nDelFail -lt 5) { Add-Log -Text "Refuse delete (too new, press may be writing): $nm" -Level ERROR }
                        $nDelFail++; continue
                    }
                }
            }
            # Emergency stop switch
            if (Test-StopRequested -Dir $scriptDir) {
                Write-Line "  ! $($script:StopFile) detected, cleanup halted" Yellow
                Add-Log -Text 'STOP: cleanup interrupted, remaining sources kept' -Level WARN
                break
            }
            $safeToDelete.Add($full)
        }

        # Size by sampling (reading every file's length is very slow)
        $mb = 0
        $sampleN = [Math]::Min(200, $safeToDelete.Count)
        $sampleSum = 0
        for ($si = 0; $si -lt $sampleN; $si++) {
            try { $sampleSum += (New-Object System.IO.FileInfo($safeToDelete[$si])).Length } catch {}
        }
        if ($sampleN -gt 0) { $mb = [long]($sampleSum / $sampleN * $safeToDelete.Count) }
        $sizeText = if ($sampleN -lt $safeToDelete.Count) { "about $([Math]::Round($mb/1MB,1)) MB (sampled)" } else { "$([Math]::Round($mb/1MB,1)) MB" }

        Write-Line "  About to delete already-merged source files: $($safeToDelete.Count) ($sizeText)" Yellow
        Write-Line '  All merged files passed the row-count check; every row is kept in SrcFile.' Gray

        $go = $false
        if ($Yes) { $go = $true }
        else {
            $ans = Read-Host '  Confirm delete? Type Y then Enter (anything else cancels)'
            if ($ans.Trim().ToUpper() -eq 'Y') { $go = $true }
        }

        if (-not $go) {
            Write-Line '  Delete cancelled. All source files kept.' Green
            Add-Log -Text 'User cancelled delete'
        }
        else {
            $estDel = [int][Math]::Round($safeToDelete.Count * 0.005 + 3)
            Write-Line "  Deleting (about $estDel s, please keep this window open)..." Gray
            $swDel = [System.Diagnostics.Stopwatch]::StartNew()
            Invoke-FastDelete -Paths $safeToDelete -StopDir $scriptDir
            $swDel.Stop()

            # Re-scan to confirm what really happened (much faster than per-file checks)
            $remain = Get-RemainingSet -Src $Src -Cutoff $cutoff
            foreach ($p in $safeToDelete) {
                if ($remain.Contains($p)) {
                    $nDelFail++
                    if ($nDelFail -le 5) { Add-Log -Text "Delete failed (kept): $p" -Level WARN }
                }
                else {
                    $nDeleted++
                    $script:Deleted.Add($p)
                }
            }
            $script:RunEnv.Deleted = $nDeleted
            Write-Line "  Deleted: $nDeleted (took $($swDel.Elapsed.TotalSeconds.ToString('N1')) s)" Green
            if ($nDelFail -gt 0) { Write-Line "  ! Delete failed: $nDelFail (kept; next run handles them)" Yellow }
            Add-Log -Text "Deleted source files: $nDeleted, failed $nDelFail, delete time $($swDel.Elapsed.TotalSeconds.ToString('N1')) s"
        }
    }
    elseif ($doDelete) {
        Write-Line '' ; Write-Line '  Nothing to clean (everything already merged).' Gray
    }

    # ---------- Summary ----------
    Write-Line ''
    Write-Line '============================================' Cyan
    Write-Line '   Done' Green
    Write-Line '============================================' Cyan
    Write-Line "  Merged files: $Out" White
    Write-Line "  Groups: $($groups.Count)   Rows added: $nNewRows   Source files deleted: $nDeleted" White
    Write-Line "  Total time: $($sw.Elapsed.TotalSeconds.ToString('N1')) s" White

    Add-Log -Text "===== DONE: $($groups.Count) groups, $nNewRows rows added, $nDeleted deleted, $($sw.Elapsed.TotalSeconds.ToString('N1')) s ====="
    Write-LogFile -ScriptDir $scriptDir

    if (-not $NoPause) { Read-Host '  Press Enter to close' }
}

# ======================== Start ========================

try {
    Main
}
catch {
    $msg = $_.Exception.Message
    $stack = ''
    if ($_.ScriptStackTrace) { $stack = $_.ScriptStackTrace }
    Write-Host ''
    Write-Host "  X Unexpected error: $msg" -ForegroundColor Red
    Write-Host "    $stack" -ForegroundColor DarkGray
    Write-Host ''

    # Must log: scheduled runs are hidden, nobody sees the screen
    $script:LogLines.Add("[ERROR] Unexpected error: $msg")
    Write-FailLog -ScriptDir (Get-ScriptDir) -Message $msg -Stack $stack
    try { Write-LogFile -ScriptDir (Get-ScriptDir) } catch {}

    if (-not $NoPause) { Read-Host '  Press Enter to close' }
    exit 1
}
finally {
    # Always release the instance lock, however we exit
    Release-RunLock -LockPath $script:LockPath
}
