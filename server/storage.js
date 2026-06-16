const dbType = process.env.DB_TYPE || "postgres";

let mod;
if (dbType === "mssql") {
  mod = await import("./storage.mssql.js");
} else {
  mod = await import("./storage.pg.js");
}

export const { DatabaseStorage, storage } = mod;
