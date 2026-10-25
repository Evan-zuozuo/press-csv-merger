#Requires -Version 5.1
# ============================================================
#  Press CSV Merger - Schedule Setup  v1.0
#
#  Registers / removes the daily automatic run using the
#  built-in Windows Task Scheduler.
#  No admin rights needed, nothing to install.
#
#  Encoding: pure ASCII (safe on any Windows language)
# ============================================================

param(
    [string]$Action,   # install | remove | status
    [string]$Time,     # HH:mm
    [switch]$NoPause
)

$script:TaskName = 'PressCSV-Merger'
$script:AutoBat  = '4_AutoRun.bat'
$script:Settings = '0_Settings_EDIT_ME.txt'

function Write-Line {
    param([string]$Text = '', [string]$Color = 'White')
    if ($Color -eq 'White') { Write-Host $Text }
    else { Write-Host $Text -ForegroundColor $Color }
}

function Get-ScriptDir {
    if ($PSScriptRoot) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return Split-Path -Parent $MyInvocation.MyCommand.Path }
    return (Get-Location).Path
}

function Test-TaskExists {
    param([string]$Name)
    try {
        $t = Get-ScheduledTask -TaskName $Name -ErrorAction Stop
        return ($null -ne $t)
    } catch {
        try {
            $null = & schtasks.exe /query /tn $Name 2>$null
            return ($LASTEXITCODE -eq 0)
        } catch { return $false }
    }
}

function Install-Task {
    param([string]$Name, [string]$BatPath, [string]$At)

    $dir = Split-Path -Parent $BatPath

    # Check settings file (otherwise scheduled runs do nothing)
    $cfgPath = Join-Path $dir $script:Settings
    $srcOk = $false
    if (Test-Path $cfgPath) {
        $raw = [System.IO.File]::ReadAllBytes($cfgPath)
        if ($raw.Length -ge 3 -and $raw[0] -eq 0xEF -and $raw[1] -eq 0xBB -and $raw[2] -eq 0xBF) {
            $cfgLines = [System.IO.File]::ReadAllLines($cfgPath, [System.Text.Encoding]::UTF8)
        } else {
            $cfgLines = [System.IO.File]::ReadAllLines($cfgPath, [System.Text.Encoding]::Default)
            if ((($cfgLines -join '') -match [char]0xFFFD)) {
                $cfgLines = [System.IO.File]::ReadAllLines($cfgPath, [System.Text.Encoding]::UTF8)
            }
        }
        foreach ($ln in $cfgLines) {
            if ($ln -match '^\s*Src\s*=\s*(.+)$') {
                $v = $Matches[1].Trim()
                if ($v -ne '') { $srcOk = $true; Write-Line "  Source folder: $v" Gray }
            }
        }
    }
    if (-not $srcOk) {
        Write-Line "  ! Warning: Src= is empty in $($script:Settings)" Yellow
        Write-Line '    Scheduled runs would exit immediately. Run "2_Merge.bat" first.' Yellow
    }

    # --- Preferred: PowerShell ScheduledTasks module ---
    try {
        Import-Module ScheduledTasks -ErrorAction Stop
        $act = New-ScheduledTaskAction -Execute $BatPath -WorkingDirectory $dir -ErrorAction Stop
        $tri = New-ScheduledTaskTrigger -Daily -At $At -ErrorAction Stop
        # -Hidden: runs in background, no window on the shop floor
        $set = New-ScheduledTaskSettingsSet -StartWhenAvailable -Hidden `
            -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Hours 6) -ErrorAction Stop
        $null = Register-ScheduledTask -TaskName $Name -Action $act -Trigger $tri `
            -Settings $set -Description 'Press CSV daily merge and clean' -Force -ErrorAction Stop
        Write-Line '  V Registered (hidden, and will catch up if missed)' Green
        return $true
    } catch {
        Write-Line ("  ! Module registration failed, trying schtasks: " + $_.Exception.Message) Yellow
    }

    # --- Fallback: schtasks.exe ---
    if (Test-TaskExists -Name $Name) {
        try { & schtasks.exe /delete /tn $Name /f | Out-Null } catch {}
    }
    $tr = '"' + $BatPath + '"'
    try { & schtasks.exe /create /tn $Name /tr $tr /sc daily /st $At /f | Out-Null } catch {}
    if (Test-TaskExists -Name $Name) { return $true }
    Write-Line '  X Registration failed (may need admin rights or policy allows it not)' Red
    return $false
}

function Remove-Task {
    param([string]$Name)
    if (-not (Test-TaskExists -Name $Name)) {
        Write-Line '  No schedule found, nothing to remove.' Gray
        return
    }
    try {
        Unregister-ScheduledTask -TaskName $Name -Confirm:$false -ErrorAction Stop
        Write-Line '  V Schedule removed' Green
        return
    } catch {
        try {
            & schtasks.exe /delete /tn $Name /f | Out-Null
            if ($LASTEXITCODE -eq 0) { Write-Line '  V Schedule removed' Green; return }
        } catch {}
        Write-Line ('  X Remove failed: ' + $_.Exception.Message) Red
    }
}

function Show-Status {
    param([string]$Name)
    if (-not (Test-TaskExists -Name $Name)) {
        Write-Line '  Status: no schedule set' Gray
        return
    }
    Write-Line '  Status: schedule is active' Green
    try {
        $t = Get-ScheduledTask -TaskName $Name -ErrorAction Stop
        $info = Get-ScheduledTaskInfo -TaskName $Name -ErrorAction Stop
        Write-Line ('    Next run : ' + $info.NextRunTime) White
        Write-Line ('    Last run : ' + $info.LastRunTime + '   result: ' + $info.LastTaskResult) White
        foreach ($tr in $t.Triggers) { Write-Line ('    Trigger  : ' + $tr.StartBoundary) Gray }
    } catch {
        & schtasks.exe /query /tn $Name /v /fo list 2>$null | Select-String -Pattern 'TaskName|Start Time|Next Run Time|Last Run Time|Last Result|Status' | ForEach-Object { Write-Line ('    ' + $_.Line.Trim()) }
    }
}

# ======================== Main ========================

$dir = Get-ScriptDir
Set-Location $dir
$bat = Join-Path $dir $script:AutoBat

Write-Line ''
Write-Line '============================================' Cyan
Write-Line '   Press CSV Merger - Schedule Setup' Cyan
Write-Line '============================================' Cyan
Write-Line ''

if ([string]::IsNullOrWhiteSpace($Action)) { $Action = 'install' }
$Action = $Action.Trim().ToLower()

if (-not (Test-Path $bat)) {
    Write-Line "  X Not found: $bat" Red
    if (-not $NoPause) { Read-Host '  Press Enter to close' }
    exit 1
}

switch ($Action) {
    'remove' {
        Write-Line '  [Remove schedule]' Yellow
        Remove-Task -Name $script:TaskName
    }
    'status' {
        Write-Line '  [Status]' Yellow
        Show-Status -Name $script:TaskName
    }
    default {
        if ([string]::IsNullOrWhiteSpace($Time)) { $Time = '08:00' }
        if ($Time -notmatch '^\d{2}:\d{2}$') { $Time = '08:00' }

        Write-Line '  [Register daily schedule]' Yellow
        Write-Line "  Run time : every day at $Time" White
        Write-Line "  Runs     : $bat" Gray
        Write-Line '  (merge + clean, background, no window, no questions)' Gray
        Write-Line ''

        if (-not $NoPause) {
            $ans = Read-Host '  Press Enter to confirm, or type another time (e.g. 07:30)'
            if (-not [string]::IsNullOrWhiteSpace($ans)) {
                if ($ans.Trim() -match '^\d{2}:\d{2}$') { $Time = $ans.Trim() }
                else { Write-Line "  ! Bad format, keeping $Time" Yellow }
            }
        }

        $ok = Install-Task -Name $script:TaskName -BatPath $bat -At $Time
        if ($ok) {
            Write-Line ''
            Write-Line '  V Done. It will run automatically every day.' Green
            Write-Line '    Check the Logs folder for results.' Gray
            Write-Line '    To remove: run "6_RemoveSchedule.bat".' Gray
        }
    }
}

Write-Line ''
if (-not $NoPause) { Read-Host '  Press Enter to close' }
