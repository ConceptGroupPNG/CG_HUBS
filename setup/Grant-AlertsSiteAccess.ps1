# =====================================================================
#  CG HUBS - give the "CG HUBS Alerts" app access to the six SharePoint
#  sites it checks, and nothing else:
#    - the five Hub sites: READ only
#    - CG_Platform (monitoring data only): FULL CONTROL, because its two
#      telemetry lists let people see only their own rows, and only Full
#      Control can see everyone's (without it the alerts see 0 visits).
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
Write-Host "This gives the app $AlertsAppId access to:" -ForegroundColor Cyan
$Sites | ForEach-Object { Write-Host "   $Host_$_   $(if ($_ -eq '/sites/CG_Platform') { 'Full Control' } else { 'Read' })" }
Read-Host "Press Enter to sign in and continue (or Ctrl+C to stop)" | Out-Null

Connect-PnPOnline -Url "$Host_/sites/CG_Platform" -Interactive -ClientId $ClientId

foreach ($s in $Sites) {
    $url = "$Host_$s"
    $existing = $null
    try { $existing = Get-PnPAzureADAppSitePermission -Site $url -AppIdentity $AlertsAppId -ErrorAction Stop } catch { $existing = $null }
    if ($existing) {
        Write-Host "  --  $s (already granted: $($existing.Roles -join ', '))" -ForegroundColor DarkYellow
    } else {
        $existing = Grant-PnPAzureADAppSitePermission -AppId $AlertsAppId -DisplayName "CG HUBS Alerts" -Site $url -Permissions Read
        Write-Host "  OK  $s - Read" -ForegroundColor Green
    }
    # Granting only offers Read/Write; Full Control is set as a second step.
    if ($s -eq "/sites/CG_Platform" -and -not ($existing.Roles -contains "fullcontrol")) {
        Set-PnPAzureADAppSitePermission -Site $url -PermissionId $existing.Id -Permissions FullControl | Out-Null
        Write-Host "  OK  $s - raised to Full Control" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "ALL DONE. Next: add the two GitHub secrets, then run the Alerts workflow." -ForegroundColor Green
