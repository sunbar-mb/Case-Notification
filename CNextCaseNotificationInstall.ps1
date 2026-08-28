<#
============================================================
 CNextCaseNotificationInstall.ps1
------------------------------------------------------------
 Sets up the "CNext Case Notification" on this PC:

   1. Creates  C:\Scripts
   2. Creates  C:\Scripts\CaseNotification
   3. Copies the required files into that folder:
        - CNextCaseNotification.ps1         (the poller)
        - CNextCaseNotificationConfig.json  (per-user settings; written interactively below)
        - CNextCaseNotificationRun.vbs          (launches the poller with no window)
        - CNextCaseNotificationIcon.ico      (toast header badge)
        - CNextCaseNotificationDocs.html     (installation & configuration guide)
        - CNextCaseNotificationToast.png    (screenshot used in the documentation)
        - CNextCaseNotificationEmail.png    (screenshot used in the documentation)
   4. Prompts for the per-user settings (organization URL, view ID, case Areas)
      and writes them to CNextCaseNotificationConfig.json. On a re-install the
      existing values are offered as defaults -- press Enter to keep them.
   5. Registers a Task Scheduler job that runs the poller every
      10 minutes, 08:00-17:00, Monday-Friday.

 Run this from the folder that contains the files above, e.g.:
     Right-click  ->  "Run with PowerShell"
   or from an elevated/normal PowerShell prompt:
     powershell -ExecutionPolicy Bypass -File .\CNextCaseNotificationInstall.ps1

 Safe to re-run / override an existing install: it stops and removes any
 previously installed version (task + running poll + old files), then installs
 fresh. The user's CNextCaseNotificationConfig.json is preserved across the
 override.
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

# --- 0. remove any existing installation first (override) ---
# Ensures a clean install over the top of an older/running version: stop the
# scheduled task, kill any poll that's currently running (so its files aren't
# locked), then clear out the old install folder. The user's config is
# preserved and re-seeded afterwards.
Write-Host "Removing any existing installation..."

# Stop and remove the existing scheduled task if present.
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    try {
        if ($existingTask.State -eq "Running") {
            Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            Write-Host "  Stopped running task: $taskName"
        }
    } catch {}
    try {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
        Write-Host "  Removed existing task: $taskName"
    } catch {
        Write-Warning "  Could not remove existing task: $_"
    }
}

# Kill any lingering poll (wscript running the .vbs, or PowerShell running the
# poller) so the old files can be replaced without file-lock errors.
$oldVbs    = Join-Path $installDir "CNextCaseNotificationRun.vbs"
$oldScript = Join-Path $installDir "CNextCaseNotification.ps1"
foreach ($procName in @("wscript", "powershell")) {
    Get-CimInstance Win32_Process -Filter "Name = '$procName.exe'" -ErrorAction SilentlyContinue |
        Where-Object {
            $_.ProcessId -ne $PID -and
            $_.CommandLine -and
            ($_.CommandLine -like "*$oldVbs*" -or $_.CommandLine -like "*$oldScript*")
        } |
        ForEach-Object {
            try {
                Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop
                Write-Host "  Stopped running poll process (PID $($_.ProcessId))."
            } catch {}
        }
}

# Preserve the user's config across the wipe, then delete the old install folder.
$configName    = "CNextCaseNotificationConfig.json"
$preservedConfig = $null
$existingConfig  = Join-Path $installDir $configName
if (Test-Path $existingConfig) {
    $preservedConfig = Join-Path $env:TEMP "$configName.installbackup"
    Copy-Item -Path $existingConfig -Destination $preservedConfig -Force
    Write-Host "  Preserved existing user config."
}
if (Test-Path $installDir) {
    try {
        Remove-Item -Path $installDir -Recurse -Force -ErrorAction Stop
        Write-Host "  Cleared old install folder: $installDir"
    } catch {
        Write-Warning "  Could not fully clear old install folder: $_"
    }
}

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

# --- 3a2. configure the user settings (interactive) ---
# Ask the user for the three per-user settings -- organization URL, view ID, and
# case Areas -- and write them into CNextCaseNotificationConfig.json, so a
# colleague never has to hand-edit the JSON. Any existing values (from a config
# preserved during an override install, else the source config shipped alongside
# this installer, else built-in placeholders) are shown as defaults: pressing
# Enter keeps them. sendEmail and notifyEmailOverride aren't prompted for here --
# they're carried over from an existing config (or left at their defaults).
$configSrc  = Join-Path $sourceDir $configName
$configDst  = Join-Path $installDir $configName

# Load whatever settings we already have as the starting point for the defaults.
$current = $null
if ($preservedConfig -and (Test-Path $preservedConfig)) {
    try { $current = Get-Content $preservedConfig -Raw | ConvertFrom-Json } catch {}
}
if (-not $current -and (Test-Path $configSrc)) {
    try { $current = Get-Content $configSrc -Raw | ConvertFrom-Json } catch {}
}

# Seed defaults from the current config where present, otherwise fall back to
# the same placeholders the shipped config uses.
$defOrgUrl    = if ($current -and $current.orgUrl)  { $current.orgUrl.ToString().Trim() }  else { "https://yourorg.crm4.dynamics.com" }
$defViewId    = if ($current -and $current.viewId)  { $current.viewId.ToString().Trim() }  else { "" }
$defAreas     = if ($current -and $current.areaFilter) { @($current.areaFilter | Where-Object { $_ -and $_.ToString().Trim() } | ForEach-Object { $_.ToString().Trim() }) } else { @() }
# Not prompted for -- preserved as-is.
$defSendEmail = if ($current -and $current.PSObject.Properties['sendEmail'] -and $null -ne $current.sendEmail) { [System.Convert]::ToBoolean($current.sendEmail) } else { $true }
$defOverride  = if ($current -and $current.PSObject.Properties['notifyEmailOverride']) { $current.notifyEmailOverride.ToString().Trim() } else { "" }

# Prompt helper: shows the current value in [brackets]; Enter keeps it.
function Read-WithDefault {
    param([string]$Prompt, [string]$Default)
    if ($Default) { $answer = Read-Host "$Prompt [$Default]" }
    else          { $answer = Read-Host $Prompt }
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
    return $answer.Trim()
}

Write-Host ""
Write-Host "Configure CNext Case Notification" -ForegroundColor Cyan
Write-Host "  Press Enter to accept the value shown in [brackets]."

# Organization URL -- strip any trailing slash so it joins cleanly with API paths.
$orgUrlIn = (Read-WithDefault "  Dynamics 365 organization URL" $defOrgUrl).TrimEnd('/')

# View ID -- the GUID after viewid= in the view's URL.
$viewIdIn = Read-WithDefault "  View ID (GUID from viewid= in the view URL)" $defViewId

# Areas -- comma-separated; blank means alert on ALL areas.
Write-Host "  Case Areas to alert on, comma-separated (e.g. Finance, Sales)."
Write-Host "  Leave blank to alert on ALL areas."
$areasIn  = Read-WithDefault "  Areas" ($defAreas -join ", ")
$areasArr = @()
if ($areasIn) {
    $areasArr = @($areasIn -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# Build the config JSON by hand so an empty area list serializes as [] (Windows
# PowerShell 5.1's ConvertTo-Json mangles empty/single-item arrays) and the
# explanatory _comment is preserved verbatim.
$jsonEsc     = { param($s) ($s -replace '\\','\\\\') -replace '"','\"' }
$areasJson   = if ($areasArr.Count -gt 0) {
    '[' + (($areasArr | ForEach-Object { '"' + (& $jsonEsc $_) + '"' }) -join ', ') + ']'
} else { '[]' }
$sendEmailJson = if ($defSendEmail) { 'true' } else { 'false' }

$configText = @"
{
  "_comment": "User settings for CNext Case Notification. orgUrl: your Dynamics 365 organization URL. viewId: GUID of the view to poll (the value after viewid= in the view URL). sendEmail: true to also send an Outlook email, false for toast-only. notifyEmailOverride: leave \"\" to alert your own mailbox (auto-detected); set an address to redirect (e.g. a shared mailbox). areaFilter: list of case Area names to alert on (e.g. \"Finance\", \"Sales\"), case-insensitive; leave [] to alert on ALL areas.",
  "orgUrl": "$(& $jsonEsc $orgUrlIn)",
  "viewId": "$(& $jsonEsc $viewIdIn)",
  "sendEmail": $sendEmailJson,
  "notifyEmailOverride": "$(& $jsonEsc $defOverride)",
  "areaFilter": $areasJson
}
"@
Set-Content -Path $configDst -Value $configText -Encoding UTF8
Write-Host "  Saved settings to: $configName" -ForegroundColor Green

# The preserved backup has now been folded into the new config; clean it up.
if ($preservedConfig -and (Test-Path $preservedConfig)) {
    Remove-Item -Path $preservedConfig -Force -ErrorAction SilentlyContinue
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

# The existing task (if any) was already removed during the cleanup step above.
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
