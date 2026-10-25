============================================================
   Press CSV Merger  V1.0   -  README
============================================================


>>>  ONLY ONE FILE YOU NEED TO EDIT  <<<

        0_Settings_EDIT_ME.txt

    Open it with Notepad, set ONE line:

        Src=D:\PressData\Export

    (the folder where the press exports its CSV files)

    That is the only setup. Everything else is automatic.


============================================================
  QUICK START  (3 steps)
============================================================

  1. Edit  0_Settings_EDIT_ME.txt   -> set Src=
  2. Double click  1_Preview.bat     -> see how many files
  3. Double click  3_MergeClean.bat  -> merge + delete sources
     (it asks you to type Y before deleting)

  Then copy the "_Merged" folder to your USB drive.


============================================================
  FILES  (all names are English on purpose - no garbled text)
============================================================

  0_Settings_EDIT_ME.txt   *** THE ONLY FILE YOU EDIT ***
  0_CheckEnv.bat           Environment check (read-only, safe)
  1_Preview.bat            Preview: count only, no write, no delete
  2_Merge.bat              Merge, keep all source files
  3_MergeClean.bat         Merge + delete sources (asks Y first)
  4_AutoRun.bat            Used by the schedule only (no questions)
  5_InstallSchedule.bat    Set up daily automatic run
  6_RemoveSchedule.bat     Remove the automatic run

  Code.ps1                 Main program (do not edit)
  Schedule.ps1             Schedule setup (do not edit)
  CheckEnv.ps1             Environment check (do not edit)
  README.txt               This file

  Logs\                    Run logs, deleted-file lists, failure reports
  _Merged\                 Merged results (created inside source folder)
  _StartupError.txt        Appears only if the program could not start
  STOP.txt                 You create this to stop the tool immediately
  _running.lock            Auto-created while running, removed after


============================================================
  WHAT IT DOES
============================================================

  The press exports 2 small CSV files every minute per machine:
    - Utility   (temperature / pressure / vacuum)
    - WORKINFO  (lot no / product / progress)

  That is 2,880 files per machine per day, 8,640 for 3 machines.

  This tool merges them into ONE file per [date + machine + type]:

      20260817_M01_Utility.csv      <- about 1440 rows
      20260817_M01_WORKINFO.csv

  Utility and WORKINFO stay SEPARATE (never mixed).
  Result: about 6 files per day instead of 8,640.

  Every row keeps a "SrcFile" column telling which original file
  it came from, so nothing loses traceability.


============================================================
  AUTOMATIC DAILY RUN
============================================================

  Prerequisite: run "2_Merge.bat" once so Src= is set.

  Then double click "5_InstallSchedule.bat":
    - default 08:00 every day
    - press Enter to accept, or type another time (e.g. 07:30)
    - runs hidden in background: no window, no questions
  To remove: "6_RemoveSchedule.bat"

  IMPORTANT: the PC must be powered on AND logged into Windows.
    If the PC is shut down or logged off, the schedule cannot run.
    Solution: keep it logged in, enable auto login, or ask IT to
    set it to "run whether user is logged on or not".

  If the PC was off at the scheduled time, it catches up on next
  login automatically.


============================================================
  SYSTEM PROTECTION  (top priority)
============================================================

  1. Process priority lowered to "below normal": it always yields
     CPU to the press program when the system is busy.
  2. Disk throttling: rests every 200 files, never saturates disk.
  3. Single instance lock: if the schedule and a manual run collide,
     the second one exits. Never two runs on the same files.
  4. Time limit (30 min default): stops safely; nothing is deleted.
  5. Emergency stop: create STOP.txt in this folder -> stops at the
     next safe point, deletes nothing.
  6. Scheduled run is hidden: operators see nothing.
  7. It only touches CSV files matching the press naming pattern,
     and only inside the source folder. No registry, no network,
     no system settings, nothing installed.


============================================================
  DATA SAFETY
============================================================

  - Merge first, verify, then delete. Deletes only after the
    row-count check passed.
  - Per-group: merge group 1 -> verify -> delete group 1 -> group 2.
    If anything interrupts, deleted ones were already safely merged
    and the rest are untouched. Just run again; it resumes.
  - Re-running is safe: already merged files are skipped, only new
    files are appended. No duplicates, no overwriting.
  - A file that fails to read is NEVER deleted.
  - Merged results are excluded from scanning, so they are never
    merged into themselves.


============================================================
  TROUBLESHOOTING
============================================================

  Q: "Nothing to merge" ?
  A: Safe window blocked it. If the press export is OFF and you want
     to include today, set SafeDays=0 in the settings file.

  Q: "Source folder not found" ?
  A: Check Src= in 0_Settings_EDIT_ME.txt. No quotes needed, even
     when the path contains spaces.

  Q: Window closes immediately ?
  A: Run from a command prompt manually to see the message.

  Q: How do I know it worked ?
  A: Look in the "Logs" folder. Every run leaves a RUNLOG file.

  Q: Something failed ?
  A: The Logs folder contains a FAILURE_*.txt report. Send that whole
     file to the engineer - it says exactly where it broke.

  Q: What can I open the merged files with ?
  A: Notepad or Excel. Standard CSV, UTF-8, comma separated.


============================================================
  ADVANCED PARAMETERS  (normally not needed)
============================================================

  Add these after "Code.ps1" inside a .bat file:

    -Src "D:\path"          Override source folder
    -Out "E:\path"          Override output (can point to USB)
    -Mode Preview           Preview mode
    -SafeDays 0             Include today's files
    -MinAgeMinutes 30       Cooling time 30 min (default 10)
    -MaxFiles 20000         Handle only 20000 this run (batch)
    -TimeoutMinutes 60      Time limit 60 min (default 30)
    -ThrottleMs 100         Rest 100 ms every 200 files (default 30)
    -DeleteSource           Delete sources after merge
    -Yes                    Do not ask for delete confirmation
    -NoPause                Do not wait for a key at the end

  Many old files (hundreds of thousands)?
    No manual batching needed. The 30-minute limit stops it safely
    and the next run continues until everything is done.
    To control it yourself: -MaxFiles 50000 or -TimeoutMinutes 120.
