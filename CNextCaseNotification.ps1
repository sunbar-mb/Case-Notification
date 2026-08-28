# ============================================================
# CNextCaseNotification.ps1
# Polls the "Unassigned Cases" view (Case/Incident) every run
# and shows a Windows toast for any case not seen before.
#
# Authenticates as your own user (no app registration needed).
# Uses the view's own FetchXML, so results always match exactly
# what you'd see opening the view in D365 CE.
# ============================================================

# --- internal constants (not user-editable) ---
# $clientId is Microsoft's pre-consented public client for Dataverse -- it needs no
# app registration in your tenant, so it stays in code rather than the config file.
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"
$seenFile = "$PSScriptRoot\CNextCaseNotificationSeen.json"

# ============================================================
# --- user configuration ---
# ============================================================
# ALL user-editable settings live in CNextCaseNotificationConfig.json (same folder
# as this script), so a colleague never has to touch the code. The values assigned
# below are only DEFAULTS -- they are used when the config file is missing, can't be
# parsed, or leaves a particular setting out. Anything present in the JSON overrides
# the matching default here.
#
#   orgUrl              Your Dynamics 365 organization URL.
#   viewId              GUID of the personal view to poll (the value after
#                       viewid= in the view's URL).
#   sendEmail           $true to also send an Outlook email; $false for toast-only.
#                       Email is sent via the Outlook desktop app (COM automation),
#                       so no credentials are stored anywhere. Requires Outlook to be
#                       installed and signed in. If Outlook's "A program is trying to
#                       send an email" prompt appears on first send, click Allow.
#   notifyEmailOverride Leave "" to send the alert to yourself -- each run detects the
#                       mailbox signed in on this PC (Exchange identity -> first
#                       Outlook account -> Dynamics sign-in; see Get-SelfEmailAddress),
#                       so the same script works for any colleague unedited. Set an
#                       address only to redirect alerts elsewhere (e.g. a shared mailbox).
#   areaFilter          Array of case Area names to alert on (e.g. "Finance", "Sales").
#                       Matching is case-insensitive. Empty = alert on all areas.
$configFile = "$PSScriptRoot\CNextCaseNotificationConfig.json"

# Defaults (used when the file or a given setting is absent).
$orgUrl              = "https://yourorg.crm4.dynamics.com"
$viewId              = "00000000-0000-0000-0000-000000000000"   # GUID of the personal view to poll (from viewid= in the view URL)
$sendEmail           = $true
$notifyEmailOverride = ""
$areaFilter          = @()

if (Test-Path $configFile) {
    try {
        $config = Get-Content $configFile -Raw | ConvertFrom-Json

        # Strings: override only when the key is present and non-blank, so an empty
        # or missing entry falls back to the default rather than blanking the setting.
        if ($config.PSObject.Properties['orgUrl'] -and $config.orgUrl.ToString().Trim()) {
            $orgUrl = $config.orgUrl.ToString().Trim()
        }
        if ($config.PSObject.Properties['viewId'] -and $config.viewId.ToString().Trim()) {
            $viewId = $config.viewId.ToString().Trim()
        }
        # notifyEmailOverride may legitimately be "" (send to self), so honour the
        # key whenever it's present, blank or not.
        if ($config.PSObject.Properties['notifyEmailOverride']) {
            $notifyEmailOverride = $config.notifyEmailOverride.ToString().Trim()
        }
        # Boolean: accept true/false (and common string forms) from the JSON.
        if ($config.PSObject.Properties['sendEmail'] -and $null -ne $config.sendEmail) {
            $sendEmail = [System.Convert]::ToBoolean($config.sendEmail)
        }
        # Area filter: keep only non-empty, trimmed entries so stray blanks in the
        # JSON don't turn into an area that can never match.
        if ($config.areaFilter) {
            $areaFilter = @($config.areaFilter | Where-Object { $_ -and $_.ToString().Trim() } | ForEach-Object { $_.ToString().Trim() })
        }
    } catch {
        Write-Warning "Could not read $configFile ($_). Using built-in defaults."
    }
}

# --- toast branding ---
# Windows shows the AppId's registered DisplayName as the toast title.
# Without this, toasts fired from powershell.exe show as "Windows PowerShell".
# Registering a custom AppId (registry-only, no Start Menu shortcut needed)
# lets us show "CNext Case Notification" instead. Safe to re-run every time.
#
# CNextCaseNotificationIcon.ico is used as the small badge next to the title
# in the header/Action Center. Falls back to the PowerShell icon if the file
# isn't present.
# NOTE: Windows caches the DisplayName per AppUserModelId, so simply changing the
# DisplayName below sometimes keeps showing the OLD toast title until reboot/re-login.
# Bumping the AppId itself forces a fresh registration, so the new title shows immediately.
$toastAppId    = "cNext.NewCaseAlert"
$toastAumidIco = "$PSScriptRoot\CNextCaseNotificationIcon.ico"
$toastIconUri  = if (Test-Path $toastAumidIco) { $toastAumidIco } else { "$PSHOME\powershell.exe,0" }

$toastRegKey = "HKCU:\Software\Classes\AppUserModelId\$toastAppId"
if (-not (Test-Path $toastRegKey)) { New-Item -Path $toastRegKey -Force | Out-Null }
Set-ItemProperty -Path $toastRegKey -Name DisplayName -Value "cNext New Case Alert" -Type String
Set-ItemProperty -Path $toastRegKey -Name IconUri -Value $toastIconUri -Type ExpandString

Import-Module MSAL.PS

# --- authenticate (disk-cached, so only the very first run prompts) ---
$redirectUri = "app://58145b91-0c36-4500-8554-080854f2ac97"   # registered redirect URI for this sample client ID

$clientApp = New-MsalClientApplication -ClientId $clientId -RedirectUri $redirectUri | Enable-MsalTokenCacheOnDisk -PassThru
$token     = $clientApp | Get-MsalToken -Scopes "$orgUrl/.default"

$headers = @{
    Authorization      = "Bearer $($token.AccessToken)"
    Accept             = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
    Prefer             = 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
}

# --- 1. pull the view's own FetchXML ---
# this is a PERSONAL view (viewType=4230 in the URL), stored in userquery, not savedquery
# (fetched fresh each run, so if you edit the view later this script stays in sync)
$viewUri  = "$orgUrl/api/data/v9.2/userqueries($viewId)?`$select=fetchxml"
$view     = Invoke-RestMethod -Uri $viewUri -Headers $headers -Method Get
$fetchXml = $view.fetchxml

# --- 2. run that FetchXML as-is against the Incident entity set ---
$encodedFetch = [System.Uri]::EscapeDataString($fetchXml)
$queryUri     = "$orgUrl/api/data/v9.2/incidents?fetchXml=$encodedFetch"

$result     = Invoke-RestMethod -Uri $queryUri -Headers $headers -Method Get
$currentIds = @($result.value.incidentid)   # @() forces array even with 0 or 1 results

# --- 2b. optional: keep only cases whose Area matches the user's filter ---
# Done BEFORE the "seen" diff so non-matching cases never enter the seen file.
# That way, if a colleague later adds an Area to their filter, cases that were
# previously skipped for that Area are treated as new and do get alerted on.
# The Area is a lookup, so its human-readable name arrives via the FormattedValue
# annotation (already requested in the Prefer header above).
if ($areaFilter.Count -gt 0 -and $currentIds.Count -gt 0) {
    $currentIds = @($currentIds | Where-Object {
        $areaUri  = "$orgUrl/api/data/v9.2/incidents($_)?`$select=_tnxt_areaid_value"
        $areaResp = Invoke-RestMethod -Uri $areaUri -Headers $headers -Method Get
        $thisArea = $areaResp.'_tnxt_areaid_value@OData.Community.Display.V1.FormattedValue'
        # -contains is case-insensitive for strings, so "finance" matches "Finance".
        $thisArea -and ($areaFilter -contains $thisArea)
    })
}

# --- 3. diff against what we've already notified on ---
$seen = if (Test-Path $seenFile) {
    @((Get-Content $seenFile -Raw | ConvertFrom-Json).ids)
} else {
    @()
}
$newIds = $currentIds | Where-Object { $_ -notin $seen }

# --- 4. for each new case, grab case code, customer, URL, and toast/email ---

# Work out who "me" is on this machine, so every colleague running this script
# gets the alert in their own mailbox without editing the file.
function Get-SelfEmailAddress {
    param($OutlookApp, $Token)

    # 1. Outlook's Exchange identity -- the primary SMTP of the mailbox the mail
    #    will actually be sent from. Correct even when the user has several
    #    aliases, and it matches the From address of the message.
    if ($OutlookApp) {
        try {
            $exchangeUser = $OutlookApp.Session.CurrentUser.AddressEntry.GetExchangeUser()
            if ($exchangeUser -and $exchangeUser.PrimarySmtpAddress) {
                return $exchangeUser.PrimarySmtpAddress
            }
        } catch { }

        # 2. Non-Exchange (IMAP/POP/shared-only) profiles have no ExchangeUser,
        #    so fall back to the SMTP address of the first configured account.
        try {
            $account = $OutlookApp.Session.Accounts | Select-Object -First 1
            if ($account -and $account.SmtpAddress) { return $account.SmtpAddress }
        } catch { }
    }

    # 3. Last resort: the account this script signed in to Dynamics with.
    #    Usually the UPN, which for most tenants equals the mail address.
    if ($Token -and $Token.Account -and $Token.Account.Username) {
        return $Token.Account.Username
    }

    return $null
}

# Connect to Outlook once (only if there's actually something to send) so a
# missing/closed Outlook doesn't block the toast notifications either.
$outlookApp = $null
if ($sendEmail -and $newIds.Count -gt 0) {
    try {
        $outlookApp = New-Object -ComObject Outlook.Application
    } catch {
        Write-Warning "Could not start Outlook for email notifications: $_"
    }
}

# Resolve the recipient once per run, not once per case.
$notifyEmail = if ($notifyEmailOverride) {
    $notifyEmailOverride
} else {
    Get-SelfEmailAddress -OutlookApp $outlookApp -Token $token
}

if ($outlookApp -and -not $notifyEmail) {
    Write-Warning "Could not determine your own email address -- skipping email notifications (toasts still work)."
}

foreach ($id in $newIds) {
    $caseUri = "$orgUrl/api/data/v9.2/incidents($id)?`$select=title,tnxt_casecode,_tnxt_accountid_value,_primarycontactid_value,_tnxt_areaid_value,description,createdon,tnxt_caseurl"
    $case    = Invoke-RestMethod -Uri $caseUri -Headers $headers -Method Get

    # Lookups (customer, contact, area) and the createdon date come back as GUIDs/raw
    # values; their human-readable text is delivered by the FormattedValue annotation
    # requested in the Prefer header above -- so read the "@...FormattedValue" sibling.
    $customerName = $case.'_tnxt_accountid_value@OData.Community.Display.V1.FormattedValue'
    $contactName  = $case.'_primarycontactid_value@OData.Community.Display.V1.FormattedValue'
    $areaName     = $case.'_tnxt_areaid_value@OData.Community.Display.V1.FormattedValue'
    $createdOn    = $case.'createdon@OData.Community.Display.V1.FormattedValue'
    $description  = $case.description
    $caseUrl      = $case.tnxt_caseurl

    # --- toast body: case identity first, then who it's for, then where it sits ---
    # BurntToast 1.x removed -AppId from every cmdlet (github.com/Windos/BurntToast
    # issue #271, "won't fix"), and its New-BTText composer wraps plain strings as
    # bindable placeholders that render with literal curly braces. So: build the
    # toast XML by hand and display it via the raw WinRT toast API -- the only
    # remaining way to get both a custom app name and clean text.
    #
    # Two constraints pull against each other:
    #   1. hint-style (text weight) is IGNORED on top-level <text>; it only works
    #      inside a <group>/<subgroup>.
    #   2. A binding with NO top-level <text> makes Windows show a fallback header
    #      ("<app name>" / "New notification") -- the junk lines we don't want.
    # Windows always renders the top-level <text> as a bold heading, so the whole
    # first line -- case number + " - " + the case Title/Description -- goes there,
    # and is bold throughout. The detail lines follow in the subgroup (where
    # hint-style is honoured). Result, no fallback header:
    #   heading  "<case number> - <case title>"  -> bold (top-level heading style)
    #   Customer / Contact / Area                 -> captionSubtle   (muted detail lines)
    $headline   = "$($case.tnxt_casecode) - $($case.title)"
    $headingXml = "<text>$([System.Security.SecurityElement]::Escape($headline))</text>"

    $bodyTextLines  = @()

    $detailLines = @()
    if ($customerName) { $detailLines += "Customer: $customerName" }
    if ($contactName)  { $detailLines += "Contact: $contactName" }
    if ($areaName)     { $detailLines += "Area: $areaName" }
    if ($detailLines.Count -gt 0) {
        $bodyTextLines += ($detailLines | ForEach-Object {
            "<text hint-style=`"captionSubtle`">$([System.Security.SecurityElement]::Escape($_))</text>"
        })
    }

    # Only emit the group when there's something to put in it (a case with no
    # customer/contact/area would otherwise produce an empty subgroup).
    $groupXml = ""
    if ($bodyTextLines.Count -gt 0) {
        $subText = $bodyTextLines -join "`n          "
        $groupXml = @"
<group>
        <subgroup>
          $subText
        </subgroup>
      </group>
"@
    }

    $imageXml = ""

    $actionsXml = ""
    if ($caseUrl) {
        $escapedUrl = [System.Security.SecurityElement]::Escape($caseUrl)
        $actionsXml = "<actions><action content=`"Open Case`" arguments=`"$escapedUrl`" activationType=`"protocol`" /></actions>"
    }

    $toastXmlString = @"
<toast>
  <visual>
    <binding template="ToastGeneric">
      $imageXml
      $headingXml
      $groupXml
    </binding>
  </visual>
  $actionsXml
</toast>
"@

    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
    [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime]                      | Out-Null

    $xmlDoc = New-Object Windows.Data.Xml.Dom.XmlDocument
    $xmlDoc.LoadXml($toastXmlString)
    $toast  = New-Object Windows.UI.Notifications.ToastNotification $xmlDoc
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($toastAppId).Show($toast)

    # --- send an email copy so the alert isn't lost if you're away from the PC ---
    if ($outlookApp -and $notifyEmail) {
        try {
            # Header: case code + title (the "what").
            # The case code itself links to the record in D365 CE (kept as a link at
            # the bottom too), so the reader can jump straight in from the header.
            $caseCodeHtml = [System.Net.WebUtility]::HtmlEncode($case.tnxt_casecode)
            if ($caseUrl) {
                $caseCodeHtml = "<a href=`"$caseUrl`">$caseCodeHtml</a>"
            }
            $bodyLines  = @("<p><b>New unassigned case: $caseCodeHtml</b></p>")

            # Key details as a compact label/value table (the "what / who / where / when").
            # Title leads the table so its bold "Title:" label lines up with the rest.
            $detailRows = @()
            if ($case.title)   { $detailRows += "<tr><td valign=`"top`"><b>Title:</b></td><td>$([System.Net.WebUtility]::HtmlEncode($case.title))</td></tr>" }
            if ($customerName) { $detailRows += "<tr><td valign=`"top`"><b>Customer:</b></td><td>$([System.Net.WebUtility]::HtmlEncode($customerName))</td></tr>" }
            if ($contactName)  { $detailRows += "<tr><td valign=`"top`"><b>Contact:</b></td><td>$([System.Net.WebUtility]::HtmlEncode($contactName))</td></tr>" }
            if ($areaName)     { $detailRows += "<tr><td valign=`"top`"><b>Area:</b></td><td>$([System.Net.WebUtility]::HtmlEncode($areaName))</td></tr>" }
            if ($createdOn)    { $detailRows += "<tr><td valign=`"top`"><b>Created on:</b></td><td>$([System.Net.WebUtility]::HtmlEncode($createdOn))</td></tr>" }
            if ($detailRows.Count -gt 0) {
                $bodyLines += "<table cellpadding=`"2`" cellspacing=`"0`" style=`"border-collapse:collapse;`">$($detailRows -join '')</table>"
            }

            # Description (the "why") -- preserve line breaks from the case text
            if ($description) {
                $descHtml   = ([System.Net.WebUtility]::HtmlEncode($description)) -replace "`r`n", "<br>" -replace "`n", "<br>"
                $bodyLines += "<p><b>Description:</b><br>$descHtml</p>"
            }

            # Action link last
            if ($caseUrl) {
                $bodyLines += "<p><a href=`"$caseUrl`">Open case in Dynamics 365</a></p>"
            }

            $mail            = $outlookApp.CreateItem(0)   # olMailItem
            $mail.To         = $notifyEmail
            $mail.Subject    = "New case: $($case.tnxt_casecode) - $($case.title)"
            $mail.HTMLBody   = ($bodyLines -join "`n")
            $mail.Send()
        } catch {
            Write-Warning "Failed to send email for case $($case.tnxt_casecode): $_"
        }
    }
}

# --- 5. persist the full current set as "seen" ---
# wrapped in an object (not a bare array) so a single-item result still
# serializes as a JSON array under Windows PowerShell 5.1
@{ ids = $currentIds } | ConvertTo-Json | Set-Content $seenFile