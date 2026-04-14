# ON-PNT® SharePoint ChatBot

A floating chat widget that answers staff questions using documents pulled directly from an on-premises SharePoint 2019 document library. Powered by Azure OpenAI GPT-4o. Embeddable in SharePoint pages via iframe.

---

## What It Does

- Connects to a SharePoint 2019 document library using NTLM/Active Directory authentication
- Syncs and indexes supported documents on demand
- Answers questions using only the content from those documents — no general knowledge
- Cites the source document (title and URL) with every answer
- Displays as a floating chat widget in the bottom-right corner of any page
- Admin settings page for configuring SharePoint connection, AI behaviour, and appearance

---

## Architecture

```
┌─────────────────────────────────────────────────────┐
│  SharePoint 2019 (on-premises)                      │
│  Document Library  ──NTLM──►  server/sharepoint.js  │
└─────────────────────────────────────────────────────┘
                                       │
                                  Sync / Ingest
                                       │
┌─────────────────────────────────────▼───────────────┐
│  Node.js / Express  (server/)                       │
│  routes.js  ──►  storage.js  ──►  PostgreSQL DB     │
│                       │                             │
│                  Azure OpenAI GPT-4o                │
│           (RAG: docs injected into prompt)          │
└──────────────────────────────┬──────────────────────┘
                               │
                          REST API
                               │
┌──────────────────────────────▼──────────────────────┐
│  React / Vite  (client/)                            │
│  Floating widget  │  Chat UI  │  Settings page      │
└─────────────────────────────────────────────────────┘
```

**Stack:**
- Frontend: React, Vite, Tailwind CSS, shadcn/ui, Framer Motion, wouter, TanStack Query
- Backend: Node.js, Express
- Database: PostgreSQL via Drizzle ORM
- AI: Azure OpenAI GPT-4o (direct REST, no SDK dependency)
- SharePoint Auth: NTLM via `httpntlm`
- Document Parsing: `mammoth` (DOCX), `pdf-parse` (PDF), native UTF-8 (TXT/CSV/MD)

---

## Pages and Routes

| Route | Description |
|---|---|
| `/` | Main page — shows the floating chat widget |
| `/settings` | Admin page — SharePoint config, AI parameters, sync |
| `/widget` | Standalone widget page for iframe embedding |

---

## Key Files

| File | Purpose |
|---|---|
| `client/src/components/ChatWidget.jsx` | Floating widget shell (minimize/maximize, trigger button) |
| `client/src/pages/Chat.jsx` | Chat UI — message list, input, markdown rendering |
| `client/src/pages/Settings.jsx` | Admin settings — appearance, AI params, SharePoint config, sync |
| `client/src/App.jsx` | Root routing |
| `server/index.js` | Express server entry point |
| `server/routes.js` | All API route handlers |
| `server/sharepoint.js` | SharePoint 2019 NTLM integration (list files, fetch content) |
| `server/storage.js` | Database access layer (all CRUD operations) |
| `server/db.js` | PostgreSQL connection and Drizzle instance |
| `shared/schema.js` | Drizzle schema — all table definitions |
| `shared/routes.js` | Shared API route definitions and Zod schemas |

---

## Database Tables

| Table | Purpose |
|---|---|
| `documents` | Synced SharePoint documents (title, content, url, source) |
| `messages` | Chat message history (role, content) |
| `sharepoint_configs` | SharePoint connection settings (single row) |
| `app_settings` | Widget appearance + AI model parameters (single row) |

---

## Environment Variables

Set these as secrets — never hardcode them in source files.

### Required — Azure OpenAI

| Variable | Description | Example |
|---|---|---|
| `AZURE_OPENAI_API_KEY` | Azure OpenAI API key | `abc123...` |
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI resource endpoint | `https://yourresource.openai.azure.com` |
| `AZURE_OPENAI_DEPLOYMENT` | Deployment / model name | `gpt-4o` |
| `AZURE_OPENAI_API_VERSION` | API version | `2024-12-01-preview` |

### Required — Database

| Variable | Description | Example |
|---|---|---|
| `DATABASE_URL` | PostgreSQL connection string | `postgresql://user:pass@host:5432/chatbot` |

### Required — Session

| Variable | Description |
|---|---|
| `SESSION_SECRET` | Random secret string for Express session signing |

> **Local fallback:** If `DATABASE_URL` is not set, the app falls back to `postgresql://postgres:pgsa@localhost:7700/chatbot`

---

## Running Locally

### Prerequisites
- Node.js 18+
- PostgreSQL running locally (or connection to a remote instance)
- Network access to your SharePoint 2019 server

### Steps

```bash
# 1. Install dependencies
npm install

# 2. Set environment variables (copy and fill in values)
cp .env.example .env   # or set them in your shell / IDE

# 3. Push the database schema
npm run db:push

# 4. Start the development server
npm run dev
```

The app runs on `http://localhost:5000` by default.
The admin settings page is at `http://localhost:5000/settings`.

---

## Running on Replit

The workflow **Start application** is pre-configured and runs:

```
NODE_ENV=development tsx server/index.js
```

Set all environment variables as **Secrets** in the Replit Secrets panel — never in `.env` files on Replit.

---

## Embedding in SharePoint 2019

### Option 1 — Content Editor Web Part (Recommended)

1. Add a **Content Editor Web Part** to the SharePoint page
2. Edit the HTML source and paste:

```html
<iframe
  src="https://your-app-url.replit.app/widget"
  style="position:fixed; bottom:0; right:0; width:420px; height:600px; border:none; z-index:9999;"
  allow="clipboard-write"
  title="ON-PNT Assistant">
</iframe>
```

3. Save and publish the page

### Option 2 — Script Editor Web Part

Paste the same `<iframe>` HTML into a Script Editor Web Part. This works on pages where Content Editor Web Parts are available but HTML editing is not.

### Notes
- The app must be hosted on an HTTPS endpoint (required for iframe embedding in modern browsers)
- For SharePoint pages with strict Content Security Policy headers, you may need to add your app's domain to the allowed frame sources in SharePoint Central Administration
- The widget is fixed-positioned so it floats over SharePoint page content

---

## Admin Settings Reference

Navigate to `/settings` to configure the following:

### Widget Appearance
| Field | Description |
|---|---|
| Assistant Name | Name shown in the chat header and tooltip |
| Welcome Message | Text shown when chat is opened with no messages |
| Not Found Message | Exact message returned when no relevant documents are found |
| Custom Instructions | Free-form instructions controlling tone, scope, persona, and rules |

### AI Model Parameters
| Field | Range | Default | Notes |
|---|---|---|---|
| Temperature | 0–2 | 0 | 0 = deterministic, 2 = creative. Neutral: 1.0 |
| Top P | 0–1 | 1 | Token sampling breadth. Neutral: 1.0 |
| Max Tokens | 100–4096 | 1500 | Maximum response length |
| Frequency Penalty | 0–2 | 0 | Reduces word repetition. Neutral: 0.0 |
| Presence Penalty | 0–2 | 0 | Encourages new topics. Neutral: 0.0 |

### SharePoint Connection
| Field | Description |
|---|---|
| Site URL | Full URL to your SharePoint site (e.g. `http://sharepoint.company.com/sites/mysite`) |
| Active Directory Domain | NetBIOS domain name (e.g. `COMPANY`) |
| Username | Service account username |
| Password | Service account password |
| Document Library Name | Name of the library to sync (e.g. `Documents`) |
| Allow self-signed certificates | Enable if SharePoint uses an internal or untrusted SSL cert |

---

## SharePoint Sync

1. Configure and save the SharePoint Connection in `/settings`
2. Click **Test Connection** to verify credentials
3. Click **Sync Document Library** to fetch and index all files
4. Sync can be run again at any time when documents are added or updated — it replaces all previously synced content

### Supported File Types

| Extension | Parser |
|---|---|
| `.txt`, `.csv`, `.md` | Native UTF-8 |
| `.docx`, `.doc` | mammoth |
| `.pdf` | pdf-parse |

> **Not yet supported:** `.xlsx`, `.pptx`, SharePoint page content

---

## Known Limitations (Phase 1)

| Limitation | Detail |
|---|---|
| No semantic / vector search | All document text is injected into the AI prompt directly. Works well for small libraries (under ~30 documents). Will degrade for large libraries. |
| No page-level citations | Citations show document title and URL only — not page numbers, section headings, or cell ranges. |
| No conversation context across turns | Each message is answered independently. Follow-up questions do not carry prior context to the AI. |
| No XLSX or PPTX support | Excel and PowerPoint files are skipped during sync. |
| No SharePoint page text | Only document library files are synced, not SharePoint wiki or publishing pages. |
| Full sync only | Every sync deletes and re-imports all documents. No incremental change detection. |
| Messages persist across sessions | Chat history is stored in the database and is shared across all users. No per-session isolation. |

---

## Planned Improvements

| Bucket | Items | Est. Effort |
|---|---|---|
| Bucket 1 — Fixes & Polish | Fix PDF parser, pass conversation history to AI, session scoping, US-only deployment | ~1 day |
| Bucket 2 — File Types & Sync | XLSX, PPTX, SharePoint page text, incremental sync | ~1 day |
| Bucket 3 — True RAG | Document chunking, Azure OpenAI embeddings, vector search, page-level citations | 1–2 weeks |
| Bucket 4 — SQL Server | Swap PostgreSQL for on-premises SQL Server (driver, schema, upserts, Windows Auth) | 1–2 days |

---

## Security Notes

- Azure OpenAI does not train on data sent via the API
- All API keys and secrets are stored in environment variables
- SharePoint credentials are stored in the database — use a service account with read-only access
- No user authentication is implemented in Phase 1
- Session-level conversation history is not yet isolated per user — do not use for sensitive data until session scoping is added (Bucket 1)
