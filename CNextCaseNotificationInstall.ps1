<#
============================================================
 CNextCaseNotificationInstall.ps1
------------------------------------------------------------
 Sets up the "CNext Case Notification" on this PC:

   1. Creates  C:\Scripts
   2. Creates  C:\Scripts\CaseNotification
   3. Copies the required files into that folder:
        - CNextCaseNotification.ps1         (the poller)
        - CNextCaseNotificationRun.vbs          (launches the poller with no window)
        - CNextCaseNotificationIcon.ico      (toast header badge)
        - CNextCaseNotificationDocs.html     (installation & configuration guide)
        - CNextCaseNotificationToast.png    (screenshot used in the documentation)
        - CNextCaseNotificationEmail.png    (screenshot used in the documentation)
   4. Registers a Task Scheduler job that runs the poller every
      10 minutes, 08:00-17:00, Monday-Friday.

 Run this from the folder that contains the files above, e.g.:
     Right-click  ->  "Run with PowerShell"
   or from an elevated/normal PowerShell prompt:
     powershell -ExecutionPolicy Bypass -File .\CNextCaseNotificationInstall.ps1

 Safe to re-run: it overwrites the copied files and replaces the
 scheduled task each time.
============================================================
#>

# --- settings ---
$rootDir    = "C:\Scripts"
$installDir = Join-Path $rootDir "CaseNotification"
$taskName   = "CNext Case Notification"
$sourceDir  = $PSScriptRoot                       # folder this installer is run from

# Files copied verbatim from the source folder. CNextCaseNotificationRun.vbs is regenerated
# (not copied) so its script path always points at the new install folder.
$filesToCopy = @(
    "CNextCaseNotification.ps1",
    "CNextCaseNotificationIcon.ico",
    "CNextCaseNotificationDocs.html",
    "CNextCaseNotificationToast.png",
    "CNextCaseNotificationEmail.png"
)

Write-Host "Installing CNext Case Notification..." -ForegroundColor Cyan

# --- 1 & 2. create folders ---
foreach ($dir in @($rootDir, $installDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Write-Host "  Created folder: $dir"
    } else {
        Write-Host "  Folder already exists: $dir"
    }
}

# --- 2b. install MSAL.PS if not already present ---
Write-Host "Checking for MSAL.PS module..."
if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
    Write-Host "  Installing MSAL.PS (this may take a moment)..."
    try {
        Install-Module -Name MSAL.PS -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "  MSAL.PS installed successfully."
    } catch {
        Write-Warning "  Could not install MSAL.PS automatically: $_"
        Write-Warning "  Install it manually with: Install-Module MSAL.PS -Scope CurrentUser"
    }
} else {
    Write-Host "  MSAL.PS already installed."
}

# --- 3. copy the required files ---
foreach ($file in $filesToCopy) {
    $src = Join-Path $sourceDir $file
    if (Test-Path $src) {
        Copy-Item -Path $src -Destination $installDir -Force
        Write-Host "  Copied: $file"
    } else {
        Write-Warning "  Missing source file (skipped): $src"
    }
}

# --- 3b. (re)generate CNextCaseNotificationRun.vbs pointing at the new location ---
# Launches PowerShell with no visible window so the 10-minute poll is silent.
$psScriptPath = Join-Path $installDir "CNextCaseNotification.ps1"
$vbsPath      = Join-Path $installDir "CNextCaseNotificationRun.vbs"
$vbsContent   = @"
Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""$psScriptPath""", 0, False
"@
Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII
Write-Host "  Wrote: CNextCaseNotificationRun.vbs (targets $psScriptPath)"

# --- 4. register the scheduled task ---
# Action: run the VBS via wscript so no console window ever flashes.
$action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$vbsPath`""

# Weekly trigger, Mon-Fri, first fire at 08:00, then repeat every 10 minutes
# for 9 hours -> last run at 17:00.
$trigger = New-ScheduledTaskTrigger -Weekly `
    -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday `
    -At 8:00AM

# Attach a repetition pattern (borrowed from a one-time trigger) to the weekly trigger.
$repetition = (New-ScheduledTaskTrigger -Once -At 8:00AM `
    -RepetitionInterval (New-TimeSpan -Minutes 10) `
    -RepetitionDuration (New-TimeSpan -Hours 9)).Repetition
$trigger.Repetition = $repetition

# Run as the currently logged-on user, only when logged on (toasts need an
# interactive session). No stored password required.
$principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5)

# Replace any existing task with the same name.
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "  Removed existing task: $taskName"
}

Register-ScheduledTask -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Polls Dynamics 365 for new unassigned cases and shows a toast + email. Runs every 10 min, 08:00-17:00, Mon-Fri." | Out-Null

Write-Host "  Registered scheduled task: $taskName"

# --- 5. run the poller once now to finish setup ---
# Doing this here (in the visible install session) means the one-time Dynamics
# sign-in happens while the user is present, instead of later from the hidden
# scheduled run where an auth dialog would be confusing or could fail silently.
# It also processes any cases already unassigned in the view right away.
Write-Host ""
Write-Host "Completing setup: running the alert once now..." -ForegroundColor Cyan
Write-Host "  You'll be asked to sign in to Dynamics (one time)." -ForegroundColor Yellow
Write-Host "  If Outlook shows an 'allow send' prompt, click Allow." -ForegroundColor Yellow
try {
    & $psScriptPath
    Write-Host "  First run complete." -ForegroundColor Green
} catch {
    Write-Warning "  First run did not complete: $_"
    Write-Warning "  The scheduled task will still run automatically on its normal schedule."
}

# --- 6. start the scheduled task once now ---
# The task otherwise sits in 'Ready' (its normal armed state) until the first
# 08:00 trigger. Starting it here proves the wscript -> vbs -> PowerShell path
# works. Sign-in is already cached and current cases are marked seen from the
# run above, so this pass is silent and near-instant.
try {
    Start-ScheduledTask -TaskName $taskName
    Write-Host "  Started the scheduled task once (it now returns to 'Ready')." -ForegroundColor Green
} catch {
    Write-Warning "  Could not start the task now: $_"
}

Write-Host ""
Write-Host "Done. Setup is complete - no further action needed." -ForegroundColor Green
Write-Host "The alert keeps running every 10 minutes between 08:00 and 17:00 on weekdays." -ForegroundColor Green
