/**
 * SQL Server table creation script — run once to set up the schema.
 * Usage: node server/migrate.mssql.js
 *
 * Equivalent of `npm run db:push` for PostgreSQL.
 * Safe to re-run — uses IF NOT EXISTS checks.
 */
import "dotenv/config";
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
     enable_chat_log     BIT NOT NULL DEFAULT 0,
      custom_theme_color  NVARCHAR(7),
     updated_at          DATETIME2 DEFAULT GETDATE()
   )`,

  `IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='knowledge_sources' AND xtype='U')
   CREATE TABLE knowledge_sources (
      id                  INT IDENTITY(1,1) PRIMARY KEY,
      name                NVARCHAR(255) NOT NULL,
      library_name        NVARCHAR(255),
      sharepoint_mode     NVARCHAR(20) NOT NULL DEFAULT 'inherit',
      description         NVARCHAR(MAX) NOT NULL DEFAULT '',
      instructions        NVARCHAR(MAX) NOT NULL DEFAULT '',
      sme_team            NVARCHAR(255) NOT NULL DEFAULT '',
      contact_method      NVARCHAR(255) NOT NULL DEFAULT '',
      contact_details     NVARCHAR(MAX) NOT NULL DEFAULT '',
      escalation_message  NVARCHAR(MAX) NOT NULL DEFAULT '',
      enabled             BIT NOT NULL DEFAULT 1,
      is_portal_wide      BIT NOT NULL DEFAULT 0,
      created_at          DATETIME2 DEFAULT GETDATE(),
      updated_at          DATETIME2 DEFAULT GETDATE()
    )`,

  `IF NOT EXISTS (
      SELECT * FROM sys.indexes
      WHERE name = 'knowledge_sources_one_portal_wide'
        AND object_id = OBJECT_ID('knowledge_sources')
    )
    CREATE UNIQUE INDEX knowledge_sources_one_portal_wide
      ON knowledge_sources(is_portal_wide)
      WHERE is_portal_wide = 1`,

  // Add enable_chat_log to existing app_settings tables (idempotent)
  `IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('app_settings') AND name = 'enable_chat_log')
   ALTER TABLE app_settings ADD enable_chat_log BIT NOT NULL DEFAULT 0`,

  `IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('app_settings') AND name = 'assistant_icon')
   ALTER TABLE app_settings ADD assistant_icon NVARCHAR(50) NOT NULL CONSTRAINT df_app_settings_assistant_icon DEFAULT 'message-circle'`,

  `IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('app_settings') AND name = 'theme')
   ALTER TABLE app_settings ADD theme NVARCHAR(50) NOT NULL CONSTRAINT df_app_settings_theme DEFAULT 'teal'`,

  `IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('app_settings') AND name = 'custom_theme_color')
   ALTER TABLE app_settings ADD custom_theme_color NVARCHAR(7) NULL`,

  `IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('knowledge_sources') AND name = 'sharepoint_mode')
   ALTER TABLE knowledge_sources ADD sharepoint_mode NVARCHAR(20) NOT NULL CONSTRAINT df_knowledge_sources_sharepoint_mode DEFAULT 'inherit'`,

  `IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('app_settings') AND name = 'launcher_label')
   ALTER TABLE app_settings ADD launcher_label NVARCHAR(40) NOT NULL CONSTRAINT df_app_settings_launcher_label DEFAULT 'Ask inSite'`,

  `IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('app_settings') AND name = 'launcher_position')
   ALTER TABLE app_settings ADD launcher_position NVARCHAR(50) NOT NULL CONSTRAINT df_app_settings_launcher_position DEFAULT 'bottom-right'`,

  `IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('app_settings') AND name = 'launcher_style')
   ALTER TABLE app_settings ADD launcher_style NVARCHAR(50) NOT NULL CONSTRAINT df_app_settings_launcher_style DEFAULT 'bubble'`,

  `IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='chat_logs' AND xtype='U')
   CREATE TABLE chat_logs (
     id                  INT IDENTITY(1,1) PRIMARY KEY,
     session_id          NVARCHAR(255) NOT NULL,
     username            NVARCHAR(255),
     user_message        NVARCHAR(MAX) NOT NULL,
     assistant_response  NVARCHAR(MAX) NOT NULL,
     created_at          DATETIME2 DEFAULT GETDATE()
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
