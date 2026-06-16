/**
 * SQL Server table creation script — run once to set up the schema.
 * Usage: node server/migrate.mssql.js
 *
 * Equivalent of `npm run db:push` for PostgreSQL.
 * Safe to re-run — uses IF NOT EXISTS checks.
 */
import { getPool, sql } from "./db.mssql.js";

const statements = [
  `IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='documents' AND xtype='U')
   CREATE TABLE documents (
     id          INT IDENTITY(1,1) PRIMARY KEY,
     title       NVARCHAR(MAX) NOT NULL,
     content     NVARCHAR(MAX) NOT NULL,
     type        NVARCHAR(255) NOT NULL,
     url         NVARCHAR(MAX) NOT NULL,
     source      NVARCHAR(255) NOT NULL DEFAULT 'manual',
     created_at  DATETIME2 DEFAULT GETDATE()
   )`,

  `IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='messages' AND xtype='U')
   CREATE TABLE messages (
     id          INT IDENTITY(1,1) PRIMARY KEY,
     role        NVARCHAR(50) NOT NULL,
     content     NVARCHAR(MAX) NOT NULL,
     created_at  DATETIME2 DEFAULT GETDATE()
   )`,

  `IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='sharepoint_configs' AND xtype='U')
   CREATE TABLE sharepoint_configs (
     id                INT IDENTITY(1,1) PRIMARY KEY,
     mode              NVARCHAR(50) NOT NULL DEFAULT 'onprem',
     site_url          NVARCHAR(MAX) NOT NULL,
     site_url_onprem   NVARCHAR(MAX),
     site_url_online   NVARCHAR(MAX),
     library_name      NVARCHAR(255) NOT NULL DEFAULT 'Documents',
     domain            NVARCHAR(255) NOT NULL DEFAULT '',
     username          NVARCHAR(255) NOT NULL DEFAULT '',
     password          NVARCHAR(MAX) NOT NULL DEFAULT '',
     allow_self_signed BIT NOT NULL DEFAULT 1,
     tenant_id         NVARCHAR(255),
     client_id         NVARCHAR(255),
     client_secret     NVARCHAR(MAX),
     last_synced_at    DATETIME2,
     updated_at        DATETIME2 DEFAULT GETDATE()
   )`,

  `IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='app_settings' AND xtype='U')
   CREATE TABLE app_settings (
     id                  INT IDENTITY(1,1) PRIMARY KEY,
     assistant_name      NVARCHAR(255) NOT NULL DEFAULT 'ON-PNT® Assistant',
     welcome_message     NVARCHAR(MAX) NOT NULL DEFAULT 'Ask me anything about your SharePoint documents.',
     not_found_message   NVARCHAR(MAX) NOT NULL DEFAULT 'I''m sorry, I couldn''t find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator.',
     custom_instructions NVARCHAR(MAX),
     temperature         FLOAT NOT NULL DEFAULT 0,
     top_p               FLOAT NOT NULL DEFAULT 1,
     max_tokens          INT NOT NULL DEFAULT 1500,
     frequency_penalty   FLOAT NOT NULL DEFAULT 0,
     presence_penalty    FLOAT NOT NULL DEFAULT 0,
     updated_at          DATETIME2 DEFAULT GETDATE()
   )`,
];

async function migrate() {
  const pool = await getPool();
  for (const stmt of statements) {
    await pool.request().query(stmt);
    const tableName = stmt.match(/name='(\w+)'/)?.[1] ?? "unknown";
    console.log(`[migrate] ✓ ${tableName}`);
  }
  console.log("[migrate] Done — all tables ready.");
  await sql.close();
}

migrate().catch((err) => {
  console.error("[migrate] FAILED:", err.message);
  process.exit(1);
});
