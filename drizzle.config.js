import { defineConfig } from "drizzle-kit";

const dbType = process.env.DB_TYPE || "postgres";

if (dbType === "mssql") {
  throw new Error(
    "[db:push] SQL Server does not use drizzle-kit.\n" +
    "Run this instead to create SQL Server tables:\n" +
    "  node server/migrate.mssql.js"
  );
}

if (!process.env.DATABASE_URL) {
  throw new Error("DATABASE_URL is not set — ensure the PostgreSQL database is provisioned.");
}

export default defineConfig({
  out: "./migrations",
  schema: "./shared/schema.js",
  dialect: "postgresql",
  dbCredentials: {
    url: process.env.DATABASE_URL,
  },
});
