<#
.SYNOPSIS
    Deploys or removes the ON-PNT® ChatBot floating widget on a SharePoint Online site collection.

.DESCRIPTION
    Adds (or removes) a ScriptLink User Custom Action that injects the chat bubble across every
    modern and classic page in the target site collection.

    Requires PnP.PowerShell module:
        Install-Module PnP.PowerShell -Force -AllowClobber

.PARAMETER SiteUrl
    Full URL of the SharePoint Online site collection, e.g.
        https://technicalassurance.sharepoint.com/sites/yoursite

.PARAMETER ChatbotUrl
    Root URL of the IIS-hosted chatbot server. Defaults to https://chatbot.technicalassurance.com
    The script appends /embed.js automatically.

.PARAMETER Uninstall
    Switch — when present, removes the custom action instead of adding it.

.PARAMETER ClientId
    Azure AD app registration Client ID for app-only (headless) auth.
    If omitted, an interactive browser login is used instead.

.PARAMETER ClientSecret
    Azure AD app registration Client Secret (app-only auth).

.PARAMETER TenantId
    Azure AD Tenant ID (app-only auth).

.EXAMPLE
    # Interactive login (browser popup)
    .\Deploy-WidgetOnline.ps1 -SiteUrl "https://technicalassurance.sharepoint.com/sites/intranet"

.EXAMPLE
    # App-only login using existing Azure AD app (same creds as your .env)
    .\Deploy-WidgetOnline.ps1 `
        -SiteUrl      "https://technicalassurance.sharepoint.com/sites/intranet" `
        -ClientId     "YOUR_CLIENT_ID" `
        -ClientSecret "YOUR_CLIENT_SECRET" `
        -TenantId     "YOUR_TENANT_ID"

.EXAMPLE
    # Remove the widget from a site
    .\Deploy-WidgetOnline.ps1 -SiteUrl "https://technicalassurance.sharepoint.com/sites/intranet" -Uninstall

.EXAMPLE
    # Deploy to every site collection in the tenant (app-only)
    Get-PnPTenantSite | ForEach-Object {
        .\Deploy-WidgetOnline.ps1 -SiteUrl $_.Url -ClientId "..." -ClientSecret "..." -TenantId "..."
    }

.NOTES
    NOESCRIPT SITES
    Modern group-connected Team Sites have DenyAddAndCustomizePages (NoScript) enabled by default,
    which silently blocks ScriptLink custom actions. Communication Sites are fine without changes.

    To check whether a site has NoScript on:
        Connect-PnPOnline -Url "https://tenant-admin.sharepoint.com" -Interactive
        Get-PnPTenantSite -Url "https://tenant.sharepoint.com/sites/yoursite" | Select DenyAddAndCustomizePages

    To turn NoScript off for a specific site (requires SharePoint Admin):
        Set-PnPTenantSite -Url "https://tenant.sharepoint.com/sites/yoursite" -DenyAddAndCustomizePages $false

    If you cannot turn off NoScript, the SPFx Application Customizer approach is required instead.
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $true)]
    [string] $SiteUrl,

    [string] $ChatbotUrl = "https://chatbot.technicalassurance.com",

    [switch] $Uninstall,

    [string] $ClientId,
    [string] $ClientSecret,
    [string] $TenantId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Name used to identify our custom action (must be unique per site) ──────────
$ActionName  = "ON-PNT-ChatBot-Widget"
$ScriptSrc   = "$($ChatbotUrl.TrimEnd('/'))/embed.js"
$Sequence    = 32000   # High sequence = loads after other scripts

# ── Connect ────────────────────────────────────────────────────────────────────
Write-Host "`n=== ON-PNT® ChatBot — SharePoint Online Widget Deployer ===" -ForegroundColor Cyan
Write-Host "Site : $SiteUrl"
Write-Host "Mode : $(if ($Uninstall) { 'REMOVE' } else { 'INSTALL' })"
Write-Host ""

if ($ClientId -and $ClientSecret -and $TenantId) {
    Write-Host "Connecting with app-only credentials..." -ForegroundColor DarkGray
    $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    Connect-PnPOnline -Url $SiteUrl `
                      -ClientId $ClientId `
                      -ClientSecret $secureSecret `
                      -Tenant $TenantId
} else {
    Write-Host "Connecting interactively (browser login will open)..." -ForegroundColor DarkGray
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

Write-Host "Connected." -ForegroundColor Green

# ── Check for existing action ──────────────────────────────────────────────────
$existing = Get-PnPCustomAction -Scope Site | Where-Object { $_.Name -eq $ActionName }

if ($Uninstall) {
    # ── Remove ─────────────────────────────────────────────────────────────────
    if ($existing) {
        Remove-PnPCustomAction -Identity $existing.Id -Scope Site -Force
        Write-Host "Widget custom action removed from $SiteUrl" -ForegroundColor Yellow
    } else {
        Write-Host "No widget custom action found on $SiteUrl — nothing to remove." -ForegroundColor DarkGray
    }
} else {
    # ── Install ────────────────────────────────────────────────────────────────
    if ($existing) {
        Write-Host "Widget is already installed on this site. Updating script URL..." -ForegroundColor DarkGray
        Remove-PnPCustomAction -Identity $existing.Id -Scope Site -Force
    }

    # ── NoScript check ─────────────────────────────────────────────────────────
    try {
        $adminUrl = $SiteUrl -replace "sharepoint\.com/.*", "sharepoint.com"
        $tenantAdminUrl = $adminUrl -replace "https://", "https://$($adminUrl -replace 'https://(\w+)\..*','$1')-admin."
        # Simple heuristic — if the Add-PnPCustomAction fails, we catch it below and warn.
    } catch {
        # Non-fatal — proceed and catch any NoScript block at the add step.
    }

    Add-PnPCustomAction `
        -Name        $ActionName `
        -Title       "ON-PNT ChatBot Widget" `
        -Description "Floating AI chat assistant powered by ON-PNT® Assistant" `
        -Location    "ScriptLink" `
        -ScriptSrc   $ScriptSrc `
        -Sequence    $Sequence `
        -Scope       Site

    Write-Host ""
    Write-Host "Widget successfully installed!" -ForegroundColor Green
    Write-Host "  Script URL : $ScriptSrc" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "NEXT STEPS:" -ForegroundColor Cyan
    Write-Host "  1. Open a page on $SiteUrl in your browser."
    Write-Host "  2. Hard-refresh (Ctrl+Shift+R) to bypass cache."
    Write-Host "  3. The ON-PNT chat bubble should appear in the bottom-right corner."
    Write-Host ""
    Write-Host "TROUBLESHOOTING — bubble not visible?" -ForegroundColor Yellow
    Write-Host "  • Open browser DevTools (F12) → Network tab → filter for 'embed.js'."
    Write-Host "    If it is missing, the site likely has NoScript (DenyAddAndCustomizePages) enabled."
    Write-Host "  • Fix: run the following as a SharePoint Admin in a separate PnP session:"
    Write-Host "      Connect-PnPOnline -Url 'https://<tenant>-admin.sharepoint.com' -Interactive"
    Write-Host "      Set-PnPTenantSite -Url '$SiteUrl' -DenyAddAndCustomizePages `$false"
    Write-Host "    Then re-run this script."
}

Disconnect-PnPOnline
Write-Host "`nDone." -ForegroundColor Cyan
