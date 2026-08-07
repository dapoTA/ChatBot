import { DatabaseStorage as PgDatabaseStorage, storage as pgStorage } from "./storage.pg.js";
import { DatabaseStorage as MssqlDatabaseStorage, storage as mssqlStorage } from "./storage.mssql.js";

const dbType = process.env.DB_TYPE || "postgres";

export const DatabaseStorage = dbType === "mssql" ? MssqlDatabaseStorage : PgDatabaseStorage;
export const storage = dbType === "mssql" ? mssqlStorage : pgStorage;
