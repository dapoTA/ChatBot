import { Pool } from "pg";
import { drizzle } from "drizzle-orm/node-postgres";
import * as schema from "../shared/schema.js";

export const pool = new Pool({
  connectionString: process.env.DATABASE_URL || "postgresql://postgres:pgsa@localhost:7700/chatbot",
});

export const db = drizzle(pool, { schema });
