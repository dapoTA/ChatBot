import { defineConfig } from "drizzle-kit";

if (!process.env.MSSQL_HOST) {
  throw new Error("MSSQL_HOST is required — ensure SQL Server connection vars are set");
}

export default defineConfig({
  out: "./migrations",
  schema: "./shared/schema.js",
  dialect: "mssql",
  dbCredentials: {
    server: process.env.MSSQL_HOST,
    port: parseInt(process.env.MSSQL_PORT || "1433"),
    database: process.env.MSSQL_DATABASE || "chatbot",
    user: process.env.MSSQL_USER,
    password: process.env.MSSQL_PASSWORD,
    trustServerCertificate: process.env.MSSQL_TRUST_CERT !== "false",
  },
});
