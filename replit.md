# ON-PNT® SharePoint ChatBot

A floating chat widget powered by OpenAI GPT that answers questions using documents synced from an on-premises SharePoint 2019 document library.

## Architecture

- **Frontend**: React + Vite + Tailwind CSS + shadcn/ui + Framer Motion + wouter
- **Backend**: Node.js + Express
- **Database**: PostgreSQL via Drizzle ORM
- **AI**: OpenAI GPT-5.1 via Replit-managed integration (RAG pattern)
- **SharePoint Auth**: NTLM (for on-premises AD) via `httpntlm`
- **Document Parsing**: `mammoth` (Word .docx), `pdf-parse` (PDF), native (plain text)

## Key Pages

- `/` — Main page with floating chat widget (bottom-right corner)
- `/settings` — SharePoint connection configuration and document sync

## Key Files

- `client/src/components/ChatWidget.tsx` — Floating chat widget with minimize/maximize
- `client/src/pages/Chat.tsx` — Chat UI component (used inside widget)
- `client/src/pages/Settings.tsx` — SharePoint settings and sync UI
- `client/src/App.tsx` — Root with nav bar and routing
- `server/routes.ts` — API routes (chat, documents, SharePoint config/sync)
- `server/sharepoint.ts` — SharePoint 2019 NTLM integration service
- `server/storage.ts` — Database access layer
- `shared/schema.ts` — Drizzle schema (documents, messages, sharepoint_configs)
- `shared/routes.ts` — Shared API route definitions

## Database Tables

- `documents` — Indexed SharePoint documents (title, content, url, source)
- `messages` — Chat message history
- `sharepoint_configs` — Single-row SharePoint connection settings

## SharePoint Integration

Uses NTLM authentication against SharePoint 2019 REST API:
- `GET /_api/web/lists/getbytitle('{library}')/items` — list files
- `GET /_api/web/GetFileByServerRelativeUrl('{path}')/$value` — fetch file content

**Supported file types:** `.docx`, `.doc`, `.pdf`, `.txt`, `.csv`, `.md`

**Deployment note:** This app must run on a server with direct network access to your SharePoint 2019 environment.

## Running

```
npm run dev        # Start development server
npm run db:push    # Push schema changes to PostgreSQL
```
