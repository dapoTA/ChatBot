# ON-PNT® SharePoint ChatBot

A floating chat widget powered by Azure OpenAI GPT-4o that answers questions using documents synced from an on-premises SharePoint 2019 document library.

## Architecture

- **Frontend**: React + Vite + Tailwind CSS + shadcn/ui + Framer Motion + wouter + TanStack Query
- **Backend**: Node.js + Express
- **Database**: PostgreSQL via Drizzle ORM
- **AI**: Azure OpenAI GPT-4o (direct REST API, credentials in env vars)
- **SharePoint Auth**: NTLM (for on-premises AD) via `httpntlm`
- **Document Parsing**: `mammoth` (DOCX), `pdf-parse` (PDF), native UTF-8 (TXT/CSV/MD)

## Key Pages

- `/` — Main page with floating chat widget (bottom-right corner)
- `/settings` — Admin settings: SharePoint connection, AI parameters, widget appearance, document sync
- `/widget` — Standalone widget page for iframe embedding in SharePoint

## Key Files

- `client/src/components/ChatWidget.jsx` — Floating chat widget with minimize/maximize
- `client/src/pages/Chat.jsx` — Chat UI component (used inside widget)
- `client/src/pages/Settings.jsx` — Admin settings page (3 forms: Appearance, AI Params, SharePoint)
- `client/src/App.jsx` — Root routing
- `server/index.js` — Express server entry point
- `server/routes.js` — All API route handlers (chat, documents, SharePoint, settings)
- `server/sharepoint.js` — SharePoint 2019 NTLM integration (list files, fetch + parse content)
- `server/storage.js` — Database access layer (all CRUD)
- `server/db.js` — PostgreSQL connection via DATABASE_URL env var
- `shared/schema.js` — Drizzle schema (documents, messages, sharepoint_configs, app_settings)
- `shared/routes.js` — Shared API route definitions and Zod validation schemas

## Database Tables

- `documents` — Synced SharePoint documents (title, content, url, source)
- `messages` — Chat message history (role, content)
- `sharepoint_configs` — SharePoint connection settings (single row)
- `app_settings` — Widget appearance + AI model parameters (single row)

## Environment Variables (Secrets)

- `AZURE_OPENAI_API_KEY` — Azure OpenAI API key
- `AZURE_OPENAI_ENDPOINT` — e.g. `https://yourresource.openai.azure.com`
- `AZURE_OPENAI_DEPLOYMENT` — Deployment name, e.g. `gpt-4o`
- `AZURE_OPENAI_API_VERSION` — e.g. `2024-12-01-preview`
- `DATABASE_URL` — PostgreSQL connection string
- `SESSION_SECRET` — Random secret for Express session signing

## SharePoint Integration

Uses NTLM authentication against SharePoint 2019 REST API:
- `GET /_api/web/lists/getbytitle('{library}')/items` — list files in library
- `GET /_api/web/GetFileByServerRelativeUrl('{path}')/$value` — fetch raw file bytes

**Supported file types:** `.docx`, `.doc`, `.pdf`, `.txt`, `.csv`, `.md`
**Not yet supported:** `.xlsx`, `.pptx`, SharePoint page content

## Running

```bash
npm run dev        # Start development server (port 5000)
npm run db:push    # Push schema changes to PostgreSQL
```

## AI Behaviour

- System prompt enforces document-only answers (`temperature: 0` by default)
- All 5 Azure OpenAI parameters configurable from `/settings` (temp, top_p, max_tokens, frequency_penalty, presence_penalty)
- Mandatory `Sources:` section with every response
- Custom admin instructions applied to every conversation
- Configurable not-found fallback message

## Known Phase 1 Limitations

- No vector/semantic search — all doc text injected into prompt directly (works for ~30 docs)
- No page-level citations — title + URL only
- No conversation context passed to AI across turns
- No XLSX/PPTX/SharePoint page support
- Full sync only (no delta/incremental)
- Messages persist across all users/sessions (no per-session isolation)

## Planned Buckets

- **Bucket 1** — PDF fix, conversation history, session scoping (~1 day)
- **Bucket 2** — XLSX, PPTX, SP pages, incremental sync (~1 day)
- **Bucket 3** — True RAG: chunking + embeddings + vector search (1–2 weeks)
- **Bucket 4** — SQL Server support: swap pg driver, rewrite schema + upserts (1–2 days)
