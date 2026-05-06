import mssql from "mssql";
import { drizzle } from "drizzle-orm/mssql-serverless";
import * as schema from "../shared/schema.js";

const config = {
  server: process.env.MSSQL_HOST || "localhost",
  port: parseInt(process.env.MSSQL_PORT || "1433"),
  database: process.env.MSSQL_DATABASE || "chatbot",
  user: process.env.MSSQL_USER || "sa",
  password: process.env.MSSQL_PASSWORD || "",
  options: {
    trustServerCertificate: process.env.MSSQL_TRUST_CERT !== "false",
  },
};

export const pool = await mssql.connect(config);
export const db = drizzle(pool, { schema });
