# =====================================================================
#  CG HUBS - create the CG_Heartbeats list on the CG_Platform site.
#
#  Background services (first: the ZKTeco bridge) update one row here
#  every few minutes, so Mission Control and the Teams alerts can tell
#  "running" from "stopped". Staff can read it but only the service can
#  change its own row; it is hidden from search.
#
#  SAFE TO RE-RUN: anything that already exists is skipped.
#  NEEDS: PowerShell 7 and PnP.PowerShell (the same set-up as before).
#  If a step shows red text, stop and send Claude a screenshot.
# =====================================================================
$ErrorActionPreference = "Stop"

$SiteUrl  = "https://conceptcompg.sharepoint.com/sites/CG_Platform"
$ClientId = "38ab6537-f037-49a3-9837-af7aab9b5ca1"   # your PnP app registration
$L        = "CG_Heartbeats"

function Ok($text)   { Write-Host "  OK  $text" -ForegroundColor Green }
function Skip($text) { Write-Host "  --  $text (already there)" -ForegroundColor DarkYellow }

Import-Module PnP.PowerShell
Read-Host "This creates the list $L on $SiteUrl. Press Enter to sign in (or Ctrl+C to stop)" | Out-Null
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$found = $null; try { $found = Get-PnPList -Identity $L -ErrorAction Stop } catch { $found = $null }
if ($found) { Skip "List $L" }
else {
    New-PnPList -Title $L -Template GenericList -OnQuickLaunch:$false | Out-Null
    Set-PnPList -Identity $L -Description "One row per background service (e.g. the ZKTeco bridge), refreshed every few minutes by the service itself. Read by Mission Control and the alerts job." | Out-Null
    Ok "List $L"
}

$fields = [ordered]@{
    LastSeen       = "<Field Type='DateTime' Name='LastSeen' StaticName='LastSeen' DisplayName='LastSeen' Format='DateTime' />"
    DeviceSerial   = "<Field Type='Text' Name='DeviceSerial' StaticName='DeviceSerial' DisplayName='DeviceSerial' MaxLength='255' />"
    DeviceLastSeen = "<Field Type='DateTime' Name='DeviceLastSeen' StaticName='DeviceLastSeen' DisplayName='DeviceLastSeen' Format='DateTime' />"
    LastSyncAt     = "<Field Type='DateTime' Name='LastSyncAt' StaticName='LastSyncAt' DisplayName='LastSyncAt' Format='DateTime' />"
    LastSyncError  = "<Field Type='Note' Name='LastSyncError' StaticName='LastSyncError' DisplayName='LastSyncError' RichText='FALSE' NumLines='6' />"
    Version        = "<Field Type='Text' Name='Version' StaticName='Version' DisplayName='Version' MaxLength='255' />"
    Host           = "<Field Type='Text' Name='Host' StaticName='Host' DisplayName='Host' MaxLength='255' />"
}
foreach ($name in $fields.Keys) {
    $f = $null; try { $f = Get-PnPField -List $L -Identity $name -ErrorAction Stop } catch { $f = $null }
    if ($f) { Skip "  column $name" }
    else { Add-PnPFieldFromXml -List $L -FieldXml $fields[$name] | Out-Null; Ok "  column $name" }
}

# Everyone may read; only the service can change its own row. Not in search.
Set-PnPList -Identity $L -ReadSecurity AllUsersReadAccess -WriteSecurity WriteOnlyMyItems -NoCrawl:$true | Out-Null
Ok "$L - read by all, changed only by the service that created the row, not in search"

Write-Host ""
Write-Host "ALL DONE. Next: update the bridge on the bridge PC (see the pull request)." -ForegroundColor Green
