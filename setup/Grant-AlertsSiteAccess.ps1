# =====================================================================
#  CG HUBS - give the "CG HUBS Alerts" app READ-ONLY access to the six
#  SharePoint sites it checks, and nothing else.
#
#  Before running: create the app in Entra with the Microsoft Graph
#  APPLICATION permission "Sites.Selected" (admin consent granted).
#  Sites.Selected on its own gives access to no sites at all - this
#  script grants each one.
#
#  Run:   .\Grant-AlertsSiteAccess.ps1 -AlertsAppId <Application (client) ID>
#  SAFE TO RE-RUN: sites already granted are skipped.
#  If a step shows red text, stop and send Claude a screenshot.
# =====================================================================
param(
    [Parameter(Mandatory = $true)][string]$AlertsAppId
)
$ErrorActionPreference = "Stop"

$Host_    = "https://conceptcompg.sharepoint.com"
$ClientId = "38ab6537-f037-49a3-9837-af7aab9b5ca1"   # your PnP app registration
$Sites    = @("/sites/training.admin", "/sites/PSP_Test", "/sites/Concept_CSA",
              "/sites/Helpdesk", "/sites/CG_Time", "/sites/CG_Platform")

Import-Module PnP.PowerShell

Write-Host ""
Write-Host "This gives the app $AlertsAppId READ access to:" -ForegroundColor Cyan
$Sites | ForEach-Object { Write-Host "   $Host_$_" }
Read-Host "Press Enter to sign in and continue (or Ctrl+C to stop)" | Out-Null

Connect-PnPOnline -Url "$Host_/sites/CG_Platform" -Interactive -ClientId $ClientId

foreach ($s in $Sites) {
    $url = "$Host_$s"
    $existing = $null
    try { $existing = Get-PnPAzureADAppSitePermission -Site $url -AppIdentity $AlertsAppId -ErrorAction Stop } catch { $existing = $null }
    if ($existing) {
        Write-Host "  --  $s (already granted: $($existing.Roles -join ', '))" -ForegroundColor DarkYellow
    } else {
        Grant-PnPAzureADAppSitePermission -AppId $AlertsAppId -DisplayName "CG HUBS Alerts" -Site $url -Permissions Read | Out-Null
        Write-Host "  OK  $s - Read" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "ALL DONE. Next: add the two GitHub secrets, then run the Alerts workflow." -ForegroundColor Green
