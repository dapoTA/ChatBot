import "dotenv/config";
import mssql from "mssql";

const config = {
  server: process.env.MSSQL_HOST || "localhost",
  port: parseInt(process.env.MSSQL_PORT || "1433"),
  database: process.env.MSSQL_DATABASE || "chatbot",
  user: process.env.MSSQL_USER || "sa",
  password: process.env.MSSQL_PASSWORD || "",
  options: { trustServerCertificate: process.env.MSSQL_TRUST_CERT !== "false" },
};

const tables = [
  {
    name: "documents",
    sql: `
      CREATE TABLE dbo.documents (
        id            INT IDENTITY(1,1) PRIMARY KEY,
        title         NVARCHAR(MAX) NOT NULL,
        content       NVARCHAR(MAX) NOT NULL,
        type          NVARCHAR(MAX) NOT NULL,
        url           NVARCHAR(MAX) NOT NULL,
        source        NVARCHAR(MAX) NOT NULL DEFAULT 'manual',
        created_at    DATETIME2     DEFAULT GETDATE()
      )`,
  },
  {
    name: "messages",
    sql: `
      CREATE TABLE dbo.messages (
        id         INT IDENTITY(1,1) PRIMARY KEY,
        role       NVARCHAR(MAX) NOT NULL,
        content    NVARCHAR(MAX) NOT NULL,
        created_at DATETIME2     DEFAULT GETDATE()
      )`,
  },
  {
    name: "sharepoint_configs",
    sql: `
      CREATE TABLE dbo.sharepoint_configs (
        id                 INT IDENTITY(1,1) PRIMARY KEY,
        site_url           NVARCHAR(MAX) NOT NULL,
        domain             NVARCHAR(MAX) NOT NULL,
        username           NVARCHAR(MAX) NOT NULL,
        password           NVARCHAR(MAX) NOT NULL,
        library_name       NVARCHAR(MAX) NOT NULL DEFAULT 'Documents',
        allow_self_signed  BIT           NOT NULL DEFAULT 1,
        last_synced_at     DATETIME2,
        updated_at         DATETIME2     DEFAULT GETDATE()
      )`,
  },
  {
    name: "app_settings",
    sql: `
      CREATE TABLE dbo.app_settings (
        id                 INT IDENTITY(1,1) PRIMARY KEY,
        assistant_name     NVARCHAR(MAX) NOT NULL DEFAULT 'ON-PNT® Assistant',
        welcome_message    NVARCHAR(MAX) NOT NULL DEFAULT 'Ask me anything about your SharePoint documents.',
        not_found_message  NVARCHAR(MAX) NOT NULL DEFAULT 'I''m sorry, I couldn''t find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator.',
        custom_instructions NVARCHAR(MAX),
        temperature        REAL          NOT NULL DEFAULT 0,
        top_p              REAL          NOT NULL DEFAULT 1,
        max_tokens         INT           NOT NULL DEFAULT 1500,
        frequency_penalty  REAL          NOT NULL DEFAULT 0,
        presence_penalty   REAL          NOT NULL DEFAULT 0,
        updated_at         DATETIME2     DEFAULT GETDATE()
      )`,
  },
];

async function main() {
  console.log("Connecting to SQL Server…");
  const pool = await mssql.connect(config);
  console.log(`Connected to ${config.database} on ${config.server}`);

  for (const table of tables) {
    const check = await pool.request().query(
      `SELECT 1 FROM sys.tables WHERE object_id = OBJECT_ID('dbo.${table.name}')`
    );
    if (check.recordset.length > 0) {
      console.log(`  [skip]   dbo.${table.name} already exists`);
    } else {
      await pool.request().query(table.sql);
      console.log(`  [created] dbo.${table.name}`);
    }
  }

  await pool.close();
  console.log("Done.");
}

main().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});
