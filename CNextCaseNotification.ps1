# ============================================================
# CNextCaseNotification.ps1
# Polls the "Unassigned Cases" view (Case/Incident) every run
# and shows a Windows toast for any case not seen before.
#
# Authenticates as your own user (no app registration needed).
# Uses the view's own FetchXML, so results always match exactly
# what you'd see opening the view in D365 CE.
# ============================================================

# --- config ---
$orgUrl   = "https://hso.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"   # Microsoft's pre-consented public client for Dataverse
$viewId   = "1a15c416-6670-eb11-a812-000d3adafcf9"   # "Unpicked cases - HSO INT Support & Optimizations" view
#$viewId   = "c01b731c-f18d-ea11-a811-000d3ab395a9"  # Testing only - your "Unassigned Cases" system view on Case (incident)
#$viewId   = "97c84bb0-7169-f111-ab0c-000d3ade82c2"  # Testing only - My Cases view - for testing
$seenFile = "$PSScriptRoot\CNextCaseNotificationSeen.json"

# --- email notification ---
# Sent via the Outlook desktop app (COM automation) so no credentials need to
# be stored here. Requires Outlook to be installed and signed in on this PC.
# If Outlook's security prompt ("A program is trying to send an email...")
# pops up on first send, click Allow (and optionally extend the allowed time)
# -- this is a one-time Outlook setting, not something the script controls.
#
# The recipient is NOT hard-coded: each run detects the mailbox of whoever is
# signed in on this PC and sends the alert to that person. So the same copy of
# this script can be handed to any colleague without editing anything.
# Resolution order (see Get-SelfEmailAddress below):
#   1. Outlook's own Exchange identity (the mailbox the mail is sent from)
#   2. The SMTP address of Outlook's first configured account
#   3. The account used to sign in to Dynamics (MSAL token)
# Only set $notifyEmailOverride if you deliberately want alerts to go somewhere
# else (e.g. a shared mailbox); leave it empty for normal "send to myself" use.
$sendEmail           = $true
$notifyEmailOverride = ""

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