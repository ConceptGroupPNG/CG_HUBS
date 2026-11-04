# =====================================================================
#  CG HUBS - set up the CG_Platform site and the two telemetry lists
#  used by Mission Control (sessions and errors).
#
#  WHAT IT DOES (it pauses before each step - press Enter to go on):
#    1. Creates the SharePoint site  /sites/CG_Platform  (skipped if it exists)
#    2. Creates the list CG_HubSessions  (one row per visit to a Hub)
#    3. Creates the list CG_HubErrors    (one row per error a user hits)
#    4. Makes both lists private: each person can only see and change
#       their OWN rows; site owners (you) see everything. Hidden from search.
#    5. Gives all Concept Group staff "Contribute" on the site, so the
#       Hubs can write their rows. No other site or list is touched.
#
#  SAFE TO RE-RUN: every step checks first and skips what already exists.
#
#  NEEDS: PowerShell 7 and PnP.PowerShell (the same set-up as before).
#  If a step shows red text, stop and send Claude a screenshot.
# =====================================================================

$ErrorActionPreference = "Stop"

$TenantName = "conceptcompg"
$SiteUrl    = "https://$TenantName.sharepoint.com/sites/CG_Platform"
$AdminUrl   = "https://$TenantName-admin.sharepoint.com"
$ClientId   = "38ab6537-f037-49a3-9837-af7aab9b5ca1"   # your PnP app registration
$Owner      = "mscully@concept.com.pg"
$TenantId   = "ea9e105c-2df6-47ef-a8e3-b4aa1d788423"
# "Everyone except external users" in this tenant
$AllStaff   = "c:0-f|rolemanager|spo-grid-all-users/$TenantId"

function Pause-Step($text) {
    Write-Host ""
    Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host $text -ForegroundColor Cyan
    Read-Host "Press Enter to continue (or Ctrl+C to stop)" | Out-Null
}
function Ok($text)   { Write-Host "  OK  $text" -ForegroundColor Green }
function Skip($text) { Write-Host "  --  $text (already there)" -ForegroundColor DarkYellow }

Import-Module PnP.PowerShell

# ---------------------------------------------------------------------
Pause-Step "STEP 1 of 5 - Create the site /sites/CG_Platform (you'll be asked to sign in)"
Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId
$site = $null
try { $site = Get-PnPTenantSite -Identity $SiteUrl -ErrorAction Stop } catch { $site = $null }
if ($site) {
    Skip "Site $SiteUrl"
} else {
    Write-Host "  Creating the site - this can take a minute or two..."
    New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title "CG Platform" -Url $SiteUrl -Owner $Owner -Wait | Out-Null
    Ok "Site created: $SiteUrl"
}

# ---------------------------------------------------------------------
Pause-Step "STEP 2 of 5 - Create the list CG_HubSessions"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

function Ensure-List($name, $description) {
    $found = $null; try { $found = Get-PnPList -Identity $name -ErrorAction Stop } catch { $found = $null }
    if ($found) { Skip "List $name" }
    else {
        New-PnPList -Title $name -Template GenericList -OnQuickLaunch:$false | Out-Null
        Set-PnPList -Identity $name -Description $description | Out-Null
        Ok "List $name"
    }
}
function Ensure-Field($list, $name, $xmlType, [switch]$Indexed) {
    $found = $null; try { $found = Get-PnPField -List $list -Identity $name -ErrorAction Stop } catch { $found = $null }
    if ($found) { Skip "  column $name"; return }
    $xml = switch ($xmlType) {
        "Text"     { "<Field Type='Text' Name='$name' StaticName='$name' DisplayName='$name' MaxLength='255' />" }
        "Note"     { "<Field Type='Note' Name='$name' StaticName='$name' DisplayName='$name' RichText='FALSE' NumLines='6' />" }
        "Number"   { "<Field Type='Number' Name='$name' StaticName='$name' DisplayName='$name' Decimals='0' />" }
        "DateTime" { "<Field Type='DateTime' Name='$name' StaticName='$name' DisplayName='$name' Format='DateTime' />" }
    }
    Add-PnPFieldFromXml -List $list -FieldXml $xml | Out-Null
    if ($Indexed) { Set-PnPField -List $list -Identity $name -Values @{ Indexed = $true } | Out-Null }
    Ok "  column $name$(if ($Indexed) { ' (indexed)' })"
}

$S = "CG_HubSessions"
Ensure-List  $S "One row per visit to a CG Hub - written by the Hubs, read by Mission Control. Rows older than 90 days are removed automatically."
Ensure-Field $S "App"        "Text"
Ensure-Field $S "UserName"   "Text"
Ensure-Field $S "UserEmail"  "Text"
Ensure-Field $S "StartedAt"  "DateTime"
Ensure-Field $S "LastSeen"   "DateTime" -Indexed
Ensure-Field $S "PageViews"  "Number"
Ensure-Field $S "Screens"    "Note"
Ensure-Field $S "ErrorCount" "Number"
Ensure-Field $S "Throttles"  "Number"
Ensure-Field $S "Build"      "Text"
Ensure-Field $S "LoadMs"     "Number"
Ensure-Field $S "Device"     "Text"
Ensure-Field $S "Browser"    "Text"

# ---------------------------------------------------------------------
Pause-Step "STEP 3 of 5 - Create the list CG_HubErrors"
$E = "CG_HubErrors"
Ensure-List  $E "One row per error a person hit in a CG Hub - written by the Hubs, read by Mission Control. Rows older than 90 days are removed automatically."
Ensure-Field $E "App"       "Text"
Ensure-Field $E "UserName"  "Text"
Ensure-Field $E "Message"   "Note"
Ensure-Field $E "Detail"    "Note"
Ensure-Field $E "SessionId" "Text"
Ensure-Field $E "Page"      "Text"
# Mission Control reads "errors since <date>", so the built-in Created column is indexed too
$created = Get-PnPField -List $E -Identity "Created"
if (-not $created.Indexed) { Set-PnPField -List $E -Identity "Created" -Values @{ Indexed = $true } | Out-Null; Ok "  column Created (indexed)" }
else { Skip "  index on Created" }

# ---------------------------------------------------------------------
Pause-Step "STEP 4 of 5 - Make both lists private (own rows only) and hide them from search"
foreach ($l in @($S, $E)) {
    # People read and change only their own rows
    Set-PnPList -Identity $l -ReadSecurity OwnReadAccess -WriteSecurity WriteOwnItems -NoCrawl:$true | Out-Null
    Ok "$l - own rows only, not in search"
}

# ---------------------------------------------------------------------
Pause-Step "STEP 5 of 5 - Let all Concept Group staff add their rows (Contribute on CG_Platform)"
Set-PnPWebPermission -User $AllStaff -AddRole "Contribute" | Out-Null
Ok "Everyone except external users - Contribute"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " ALL DONE. CG_Platform, CG_HubSessions and CG_HubErrors are ready." -ForegroundColor Green
Write-Host " Tell Claude it finished - the Hubs start writing once the" -ForegroundColor Green
Write-Host " telemetry pull request is merged." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
