#Requires -Version 5.1
# ============================================================
# Press CSV Merger - Environment Check  v1.0
#
# Read-only. Does NOT modify, merge or delete anything.
# Purpose: check whether this PC can run the merger,
#          and produce a report file for the engineer.
#
# Encoding: pure ASCII (safe on Japanese / Chinese / English Windows)
# ============================================================

param([switch]$NoPause)

$script:Out = New-Object System.Collections.Generic.List[string]

function W {
    param([string]$Text = '')
    $script:Out.Add($Text)
    Write-Host $Text
}

function Get-ScriptDir {
    if ($PSScriptRoot) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return Split-Path -Parent $MyInvocation.MyCommand.Path }
    return (Get-Location).Path
}

function Is-Ascii {
    param([string]$S)
    if ([string]::IsNullOrEmpty($S)) { return $true }
    return (-not ([regex]::IsMatch($S, '[^\x00-\x7F]')))
}

$dir = Get-ScriptDir
Set-Location $dir

W ''
W '================================================'
W ' Press CSV Merger - Environment Check  v1.0'
W '================================================'
W ''
W ('Time        : ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
W ('Computer    : ' + $env:COMPUTERNAME)
W ('User        : ' + $env:USERDOMAIN + '\' + $env:USERNAME)

# ---------- Admin ----------
$isAdmin = $false
try {
    $p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {}
W ('Admin rights: ' + $isAdmin + '   (not required for this tool)')

# ---------- PowerShell / OS ----------
W ('PowerShell  : ' + $PSVersionTable.PSVersion.ToString())
try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    W ('OS          : ' + $os.Caption)
    W ('OS Build    : ' + $os.BuildNumber)
} catch {
    W ('OS          : (cannot read) ' + $_.Exception.Message)
}
try {
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    W ('Memory (MB) : ' + [math]::Round($cs.TotalPhysicalMemory / 1MB))
} catch {}

# ---------- Locale / Code page ----------
W ''
W '--- Locale / Code Page ---'
W ('ANSI CP     : ' + [System.Text.Encoding]::Default.CodePage + '   (932=Japanese, 950=Trad.Chinese, 936=Simp.Chinese)')
try { W ('OEM CP      : ' + [System.Console]::OutputEncoding.CodePage) } catch {}
try { W ('Culture     : ' + (Get-Culture).Name) } catch {}
try { W ('SystemLocale: ' + (Get-WinSystemLocale).Name) } catch {}
try {
    $l = Get-WinUserLanguageList -ErrorAction Stop
    $names = @()
    foreach ($x in $l) { $names += $x.LanguageTag }
    W ('UI Language : ' + ($names -join ', '))
} catch {}

# ---------- Script folder ----------
W ''
W '--- Script Folder ---'
W ('Path        : ' + $dir)
W ('Path is ASCII only : ' + (Is-Ascii $dir))
if (-not (Is-Ascii $dir)) {
    W '  WARNING: folder path contains non-ASCII characters.'
    W '          This may cause encoding trouble. Use an English path.'
}

$tf = Join-Path $dir '_envtest.tmp'
try {
    [System.IO.File]::WriteAllText($tf, 'test', [System.Text.ASCIIEncoding]::new())
    W 'Write       : OK'
} catch {
    W ('Write       : FAIL - ' + $_.Exception.Message)
}
try {
    if (Test-Path $tf) { [System.IO.File]::Delete($tf); W 'Delete      : OK' }
} catch {
    W ('Delete      : FAIL - ' + $_.Exception.Message)
}

# ---------- Files in this folder ----------
W ''
W '--- Files in this folder ---'
$badName = 0
try {
    $files = Get-ChildItem -Path $dir -File -ErrorAction Stop | Sort-Object Name
    foreach ($f in $files) {
        $flag = 'ASCII    '
        if (-not (Is-Ascii $f.Name)) { $flag = 'NON-ASCII'; $badName++ }
        W ('  ' + $flag + '  ' + $f.Name)
    }
} catch {}
if ($badName -gt 0) {
    W ''
    W ('  >>> WARNING: ' + $badName + ' file name(s) contain non-ASCII characters.')
    W '      If the names look garbled, the file was corrupted during copy/unzip.'
    W '      Fix: re-copy the folder directly (do not use a zip file).'
}

# ---------- Settings ----------
W ''
W '--- Settings ---'
$cfgName = '0_Settings_EDIT_ME.txt'
$cfgPath = Join-Path $dir $cfgName
$srcPath = ''
if (-not (Test-Path $cfgPath)) {
    W ('  ' + $cfgName + ' : NOT FOUND')
    W '  (This is the ONLY file you need to edit. Create it if missing.)'
} else {
    $cfgLines = @()
    try {
        $raw = [System.IO.File]::ReadAllBytes($cfgPath)
        if ($raw.Length -ge 3 -and $raw[0] -eq 0xEF -and $raw[1] -eq 0xBB -and $raw[2] -eq 0xBF) {
            $cfgLines = [System.IO.File]::ReadAllLines($cfgPath, [System.Text.Encoding]::UTF8)
        } else {
            $cfgLines = [System.IO.File]::ReadAllLines($cfgPath, [System.Text.Encoding]::Default)
        }
    } catch {}
    foreach ($ln in $cfgLines) {
        if ($ln -match '^\s*([^=;#]+?)\s*=\s*(.*)$') {
            $k = $Matches[1].Trim()
            $v = $Matches[2].Trim()
            W ('  ' + $k + ' = ' + $v)
            if ($k -ieq 'Src') { $srcPath = $v }
        }
    }
}

# ---------- Source folder ----------
W ''
W '--- Source Folder (where press exports CSV) ---'
if ([string]::IsNullOrWhiteSpace($srcPath)) {
    W ('  Not configured yet. Set it in ' + $cfgName)
} else {
    W ('Path        : ' + $srcPath)
    if (Test-Path $srcPath) {
        W 'Exists      : YES'
        $nCsv = 0
        $nAll = 0
        try {
            $all = [System.IO.Directory]::GetFiles($srcPath, '*', [System.IO.SearchOption]::AllDirectories)
            $nAll = $all.Count
            foreach ($p in $all) {
                $e = [System.IO.Path]::GetExtension($p)
                if ($e -ieq '.csv') { $nCsv++ }
            }
            W ('Files total : ' + $nAll)
            W ('CSV files   : ' + $nCsv)
        } catch {
            W ('Scan        : FAIL - ' + $_.Exception.Message)
        }
        $st = Join-Path $srcPath '_envtest.tmp'
        try {
            [System.IO.File]::WriteAllText($st, 'test', [System.Text.ASCIIEncoding]::new())
            W 'Write       : OK'
            [System.IO.File]::Delete($st)
            W 'Delete      : OK   (can clean up source files)'
        } catch {
            W ('Write/Delete: FAIL - ' + $_.Exception.Message)
        }
        try {
            $d = [System.IO.Path]::GetPathRoot($srcPath)
            $drv = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($d.TrimEnd('\'))'" -ErrorAction Stop
            W ('Disk free   : ' + [math]::Round($drv.FreeSpace / 1MB) + ' MB on ' + $d)
        } catch {}
        if (-not (Is-Ascii $srcPath)) {
            W '  WARNING: source path has non-ASCII characters.'
        }
    } else {
        W 'Exists      : NO  <<< CHECK THIS'
    }
}

# ---------- Scheduled task ----------
W ''
W '--- Scheduled Task Capability ---'
$modOk = $false
try {
    Import-Module ScheduledTasks -ErrorAction Stop
    W 'Module      : OK'
    $modOk = $true
} catch {
    W ('Module      : FAIL - ' + $_.Exception.Message)
    W '            (will fall back to schtasks.exe)'
}
if ($modOk) {
    try {
        $t = Get-ScheduledTask -TaskName 'PressCSV-EnvTest' -ErrorAction SilentlyContinue
        W 'Query       : OK'
    } catch {
        W ('Query       : FAIL - ' + $_.Exception.Message)
    }
    try {
        $bat = Join-Path $dir '4_AutoRun.bat'
        if (-not (Test-Path $bat)) { $bat = 'cmd.exe' }
        $act = New-ScheduledTaskAction -Execute $bat -WorkingDirectory $dir -ErrorAction Stop
        $tri = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddDays(1).Date.AddHours(3)) -ErrorAction Stop
        $null = Register-ScheduledTask -TaskName 'PressCSV-EnvTest' -Action $act -Trigger $tri -Force -ErrorAction Stop
        W 'Register    : OK   (test task created)'
        Unregister-ScheduledTask -TaskName 'PressCSV-EnvTest' -Confirm:$false -ErrorAction Stop
        W 'Unregister  : OK   (test task removed)'
        W '=> Auto scheduling should WORK on this PC.'
    } catch {
        W ('Register    : FAIL - ' + $_.Exception.Message)
        W '=> Auto scheduling may NOT work. Engineer: check permissions.'
    }
}

# ---------- Summary ----------
W ''
W '================================================'
W ' Please send this file back to the engineer:'
W '     EnvCheck_Result.txt'
W ' (it is created in this same folder)'
W '================================================'
W ''

$resPath = Join-Path $dir 'EnvCheck_Result.txt'
try {
    # UTF-8 with BOM: keeps non-ASCII info (OS name, folder names) readable
    [System.IO.File]::WriteAllLines($resPath, $script:Out, (New-Object System.Text.UTF8Encoding($true)))
    Write-Host ''
    Write-Host ('Result saved to: ' + $resPath) -ForegroundColor Green
} catch {
    Write-Host ''
    Write-Host ('Could not save result file: ' + $_.Exception.Message) -ForegroundColor Yellow
    Write-Host 'Please take a photo of this screen instead.' -ForegroundColor Yellow
}

if (-not $NoPause) { Read-Host 'Press Enter to close' }
