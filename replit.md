# ON-PNT® SharePoint ChatBot

A floating chat widget powered by Azure OpenAI GPT-4o that answers questions using documents synced from an on-premises SharePoint 2019 document library.

## Architecture

- **Frontend**: React + Vite + Tailwind CSS + shadcn/ui + Framer Motion + wouter + TanStack Query
- **Backend**: Node.js + Express
- **Database**: PostgreSQL (default) or SQL Server — toggled via `DB_TYPE` env var
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
- `server/storage.js` — Dynamic re-exporter: picks storage.pg.js or storage.mssql.js based on DB_TYPE
- `server/storage.pg.js` — PostgreSQL storage (Drizzle ORM)
- `server/storage.mssql.js` — SQL Server storage (raw mssql queries)
- `server/db.js` — PostgreSQL connection via DATABASE_URL env var (Drizzle)
- `server/db.mssql.js` — SQL Server connection pool (mssql package)
- `server/migrate.mssql.js` — One-time SQL Server table creation script (replaces db:push for mssql)
- `server/crypto.js` — AES-256-GCM encryption for passwords stored in DB
- `shared/schema.js` — Drizzle schema (PostgreSQL only)
- `shared/routes.js` — Shared API route definitions and Zod validation schemas

## Database Tables

- `documents` — Synced SharePoint documents (title, content, url, source)
- `messages` — Chat message history (role, content)
- `sharepoint_configs` — SharePoint connection settings (single row)
- `app_settings` — Widget appearance + AI model parameters (single row)

## Environment Variables (Secrets)

### Always required
- `AZURE_OPENAI_API_KEY` — Azure OpenAI API key
- `AZURE_OPENAI_ENDPOINT` — e.g. `https://yourresource.openai.azure.com`
- `AZURE_OPENAI_DEPLOYMENT` — Deployment name, e.g. `gpt-4o`
- `AZURE_OPENAI_API_VERSION` — e.g. `2024-12-01-preview`
- `SESSION_SECRET` — Random secret for Express session signing
- `ENCRYPTION_KEY` — 64-char hex key for AES-256 password encryption
- `DB_TYPE` — `postgres` (default) or `mssql`

### PostgreSQL (DB_TYPE=postgres or unset)
- `DATABASE_URL` — PostgreSQL connection string, e.g. `postgresql://user:pass@host:5432/dbname`

### SQL Server (DB_TYPE=mssql)
- `MSSQL_SERVER` — SQL Server hostname or IP
- `MSSQL_DATABASE` — Database name, e.g. `chatbot`
- `MSSQL_USER` — SQL Server login
- `MSSQL_PASSWORD` — SQL Server password
- `MSSQL_PORT` — Port (default `1433`)
- `MSSQL_TRUST_CERT` — Set to `false` to enforce certificate validation (default `true`)
- `MSSQL_ENCRYPT` — Set to `true` to force TLS encryption (default `false`)

### SharePoint Online (optional)
- `SHAREPOINT_TENANT_ID` — Azure AD Tenant ID
- `SHAREPOINT_CLIENT_ID` — App registration Client ID
- `SHAREPOINT_CLIENT_SECRET` — App registration Client Secret

### Administration access on IIS
- `ADMIN_USERS` — Comma-, semicolon-, or newline-separated Windows usernames allowed to manage knowledge sources. Required in production; IIS Windows Authentication must promote `AUTH_USER` through iisnode.

## Running

```bash
npm run dev        # Start development server (port 5000)

# PostgreSQL (default):
npm run db:push    # Push schema changes to PostgreSQL

# SQL Server:
node server/migrate.mssql.js   # Create SQL Server tables (run once)
```

## Database Toggle

Set `DB_TYPE` in your `.env` to switch databases:

| `DB_TYPE` | Driver | Schema tool |
|---|---|---|
| `postgres` (default) | `pg` + Drizzle ORM | `npm run db:push` |
| `mssql` | `mssql` raw SQL | `node server/migrate.mssql.js` |

No code changes needed — just change the env var and restart.

## AI Behaviour

- System prompt enforces document-only answers (`temperature: 0` by default)
- All 5 Azure OpenAI parameters configurable from `/settings` (temp, top_p, max_tokens, frequency_penalty, presence_penalty)
- Mandatory `Sources:` section with every response
- Custom admin instructions applied to every conversation
- Configurable not-found fallback message

## IIS Deployment (Windows Server)

1. Install Node.js v22 LTS on the server
2. Install **iisnode** and **URL Rewrite** IIS modules
3. Run `npm run build` to produce `dist/public/` (static frontend)
4. Copy project to IIS site folder
5. Add a `web.config` pointing iisnode at `server/index.js`
6. Set all env vars in IIS Application Pool → Advanced Settings → Environment Variables
7. Run `node server/migrate.mssql.js` once to create SQL Server tables

## User Preferences

- **Never mention credits, credit balance, or suggest the user check their credits.** Do not raise this topic under any circumstances.

## Known Phase 1 Limitations

- No vector/semantic search — all doc text injected into prompt directly (works for ~30 docs)
- No page-level citations — title + URL only
- No conversation context passed to AI across turns
- No XLSX/PPTX/SharePoint page support
- Full sync only (no delta/incremental)
- Messages persist across all users/sessions (no per-session isolation)
