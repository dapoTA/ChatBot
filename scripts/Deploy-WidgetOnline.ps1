<#
.SYNOPSIS
    Deploys or removes the ON-PNT ChatBot floating widget on a SharePoint Online site collection.

.DESCRIPTION
    Adds (or removes) a ScriptLink User Custom Action that injects the chat bubble across every
    page in the target site collection.

    Requires PnP.PowerShell module:
        Install-Module PnP.PowerShell -Force -AllowClobber

.PARAMETER SiteUrl
    Full URL of the SharePoint Online site collection, e.g.
        https://technicalassurance.sharepoint.com/sites/yoursite

.PARAMETER ChatbotUrl
    Root URL of the IIS-hosted chatbot server.
    Defaults to https://chatbot.technicalassurance.com
    The script appends /embed.js automatically.

.PARAMETER Uninstall
    Switch -- when present, removes the custom action instead of adding it.

.PARAMETER ClientId
    Azure AD app registration Client ID for app-only (headless) auth.
    If omitted, an interactive browser login is used instead.

.PARAMETER ClientSecret
    Azure AD app registration Client Secret (app-only auth).

.PARAMETER TenantId
    Azure AD Tenant ID (app-only auth).

.EXAMPLE
    # Interactive login (browser popup)
    .\Deploy-WidgetOnline.ps1 -SiteUrl "https://technicalassurance.sharepoint.com/sites/insite2"

.EXAMPLE
    # App-only login using Azure AD app credentials
    .\Deploy-WidgetOnline.ps1 `
        -SiteUrl      "https://technicalassurance.sharepoint.com/sites/insite2" `
        -ClientId     "YOUR_CLIENT_ID" `
        -ClientSecret "YOUR_CLIENT_SECRET" `
        -TenantId     "YOUR_TENANT_ID"

.EXAMPLE
    # Remove the widget from a site
    .\Deploy-WidgetOnline.ps1 -SiteUrl "https://technicalassurance.sharepoint.com/sites/insite2" -Uninstall

.NOTES
    NOSCRIPT SITES
    Modern group-connected Team Sites have NoScript (DenyAddAndCustomizePages) enabled by default,
    which silently blocks ScriptLink custom actions. Communication Sites are fine.

    To turn NoScript off for a site (requires SharePoint Admin):
        Connect-PnPOnline -Url "https://technicalassurance-admin.sharepoint.com" -Interactive
        Set-PnPTenantSite -Url "https://technicalassurance.sharepoint.com/sites/yoursite" -DenyAddAndCustomizePages $false
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

$ActionName = "ON-PNT-ChatBot-Widget"
$ScriptSrc  = ($ChatbotUrl.TrimEnd('/')) + "/embed.js"
$Sequence   = 32000

Write-Host ""
Write-Host "=== ON-PNT ChatBot -- SharePoint Online Widget Deployer ===" -ForegroundColor Cyan
Write-Host "Site : $SiteUrl"
Write-Host "Mode : $(if ($Uninstall) { 'REMOVE' } else { 'INSTALL' })"
Write-Host ""

if ($ClientId -and $ClientSecret -and $TenantId) {
    Write-Host "Connecting with app-only credentials..." -ForegroundColor DarkGray
    $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -ClientSecret $secureSecret -Tenant $TenantId
} else {
    Write-Host "Connecting interactively (browser login will open)..." -ForegroundColor DarkGray
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

Write-Host "Connected." -ForegroundColor Green

$existing = Get-PnPCustomAction -Scope Site | Where-Object { $_.Name -eq $ActionName }

if ($Uninstall) {
    if ($existing) {
        Remove-PnPCustomAction -Identity $existing.Id -Scope Site -Force
        Write-Host "Widget custom action removed from $SiteUrl" -ForegroundColor Yellow
    } else {
        Write-Host "No widget custom action found on this site -- nothing to remove." -ForegroundColor DarkGray
    }
} else {
    if ($existing) {
        Write-Host "Widget already installed -- updating script URL..." -ForegroundColor DarkGray
        Remove-PnPCustomAction -Identity $existing.Id -Scope Site -Force
    }

    Add-PnPCustomAction `
        -Name        $ActionName `
        -Title       "ON-PNT ChatBot Widget" `
        -Description "Floating AI chat assistant powered by ON-PNT Assistant" `
        -Location    "ScriptLink" `
        -ScriptSrc   $ScriptSrc `
        -Sequence    $Sequence `
        -Scope       Site

    Write-Host ""
    Write-Host "Widget successfully installed!" -ForegroundColor Green
    Write-Host "  Script URL : $ScriptSrc" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "NEXT STEPS:" -ForegroundColor Cyan
    Write-Host "  1. Open any page on $SiteUrl in your browser."
    Write-Host "  2. Hard-refresh (Ctrl+Shift+R) to bypass cache."
    Write-Host "  3. The chat bubble should appear in the bottom-right corner."
    Write-Host ""
    Write-Host "Bubble not visible?" -ForegroundColor Yellow
    Write-Host "  Open DevTools (F12) > Network > filter for 'embed.js'."
    Write-Host "  If embed.js is missing, the site has NoScript enabled. Fix with:"
    Write-Host "    Connect-PnPOnline -Url 'https://technicalassurance-admin.sharepoint.com' -Interactive"
    Write-Host "    Set-PnPTenantSite -Url '$SiteUrl' -DenyAddAndCustomizePages $false"
    Write-Host "  Then re-run this script."
}

Disconnect-PnPOnline
Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
