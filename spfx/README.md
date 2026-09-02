# ON-PNT ChatBot — SPFx Web Part and Application Customizer

This SharePoint Framework solution contains two complementary experiences:

- **inSite Assistant Web Part** — a page-level, source-aware chat card for modern SharePoint Online pages.
- **ChatBot Widget Application Customizer** — the existing floating chat bubble available across modern pages.

The Web Part is additional. Installing it does not replace or remove the floating widget.

---

## Prerequisites

| Requirement | Version |
|---|---|
| Node.js | 18 LTS (18.17.1 or later, but < 19) |
| npm | 9+ (bundled with Node 18) |

> **Important:** SPFx 1.18.2 requires Node.js 18 specifically. Run `node -v` to confirm.
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

### Step 2 — Add the page-level Web Part

1. Open the modern SharePoint Online page where the assistant should appear.
2. Select **Edit**.
3. Select the **+** Web Part picker.
4. Search for **inSite Assistant** and add it to the page.
5. Edit the Web Part properties if the chatbot server is not
   `https://chatbot.technicalassurance.com`.
6. Publish or republish the page.

The Web Part uses the signed-in SharePoint user's login name for chat logging and
loads enabled knowledge sources from the shared chatbot API. Questions include the
selected source ID, so PTO, HR, Company Policies, and All Portal Sources use the
same filtering behavior as the hosted floating widget.

### Step 3 — Provision or retain the floating customizer

The tenant-wide package setting makes the solution available, but it does not by
itself activate the Application Customizer custom action. Provision the custom
action at tenant or site scope, or retain the existing custom action if the
floating widget is already deployed.

Use component ID
`a3b4c5d6-e7f8-4012-abcd-ef1234567890` and location
`ClientSideExtension.ApplicationCustomizer`.

After the custom action is present, the chat bubble should appear on modern
SharePoint Online pages within a few minutes. Hard-refresh (Ctrl+Shift+R) if it
doesn't appear immediately.

---

## Update the chatbot server URL

The URL defaults to `https://chatbot.technicalassurance.com` (hardcoded in the manifest and TypeScript).

To change it for the Web Part, edit the Web Part and update **Chatbot server URL**
in its property pane.

To change it for the tenant-wide floating widget:
1. Edit `DEFAULT_CHATBOT_URL` in `src/extensions/chatbotWidget/ChatbotWidgetApplicationCustomizer.ts`.
2. Rebuild and re-upload the `.sppkg`.

---

## Remove the solution

1. Go to the App Catalog
2. Find `chatbot-widget-client-side-solution`
3. Click **...** → **Delete**

Deleting the solution removes both the Web Part and the Application Customizer.
To remove only the page-level experience, remove the Web Part from each page and
republish it.

---

## API and CORS requirements

The chatbot host must expose:

- `GET /api/knowledge-sources/options`
- `POST /api/chat`

The chatbot server must allow requests from SharePoint Online. This project
already enables CORS for those endpoints. The Web Part sends `sourceId`,
`sessionId`, and the SharePoint login name with every question.

The login name is a client-supplied identity hint used by the existing shared
chat contract; it is not a server-verified audit identity. Use an authenticated
proxy or signed identity assertion before treating chat-log usernames as
security evidence.

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
