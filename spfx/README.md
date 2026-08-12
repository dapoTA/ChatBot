# ON-PNT ChatBot Widget — SPFx Application Customizer

Injects the floating ON-PNT chat bubble on every SharePoint Online modern page via a SharePoint Framework Application Customizer. No custom script permissions required.

---

## Prerequisites

| Requirement | Version |
|---|---|
| Node.js | 18 LTS (18.17.1 or later, but < 19) |
| npm | 9+ (bundled with Node 18) |

> **Important:** SPFx 1.20 requires Node.js 18 specifically. Run `node -v` to confirm.
> If you have a different version, install Node 18 from https://nodejs.org/en/download (use the LTS installer).

---

## Build

```powershell
# From the repo root
cd spfx
npm install          # first time only — takes 2-3 minutes
npm run dist         # bundles and packages in one step
```

The output package is at:
```
spfx/sharepoint/solution/chatbot-widget.sppkg
```

---

## Deploy to SharePoint Online

### Step 1 — Upload to App Catalog

1. Go to your **SharePoint Admin Center**: `https://technicalassurance-admin.sharepoint.com`
2. Left menu → **More features** → **Apps** → **Open**
3. Click **App Catalog** (or create one if prompted)
4. Click **Distribute apps for SharePoint** → **Upload**
5. Upload `spfx/sharepoint/solution/chatbot-widget.sppkg`
6. When prompted: check **"Make this solution available to all sites in the organization"**
7. Click **Deploy**

### Step 2 — Confirm deployment

The chat bubble should appear on SharePoint Online modern pages within a few minutes. Hard-refresh (Ctrl+Shift+R) if it doesn't appear immediately.

---

## Update the chatbot server URL

The URL defaults to `https://chatbot.technicalassurance.com` (hardcoded in the manifest and TypeScript).

To change it:
1. Edit `DEFAULT_CHATBOT_URL` in `src/extensions/chatbotWidget/ChatbotWidgetApplicationCustomizer.ts`
2. Rebuild and re-upload the `.sppkg`

---

## Remove the widget

1. Go to the App Catalog
2. Find `chatbot-widget-client-side-solution`
3. Click **...** → **Delete**

---

## Security note

This solution has `safeWithCustomScriptDisabled: true` and `requiresCustomScript: false`, meaning it works with NoScript enabled. Do **not** disable `DenyAddAndCustomizePages` for this solution — it is not required.

If you previously disabled NoScript to test the User Custom Action approach, re-enable it with:

```powershell
# Connect to SPO (Windows PowerShell, as admin)
Connect-SPOService -Url "https://technicalassurance-admin.sharepoint.com"

# Re-enable NoScript on the site
Set-SPOSite -Identity "https://technicalassurance.sharepoint.com/sites/insite2" -DenyAddAndCustomizePages 1

# Confirm tenant level was never changed (should show blank or Enabled)
Get-SPOTenant | Select DenyAddAndCustomizePages
```
