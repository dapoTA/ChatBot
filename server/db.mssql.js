import sql from "mssql";

const config = {
  server: process.env.MSSQL_SERVER || "localhost",
  database: process.env.MSSQL_DATABASE || "chatbot",
  user: process.env.MSSQL_USER,
  password: process.env.MSSQL_PASSWORD,
  port: parseInt(process.env.MSSQL_PORT || "1433"),
  options: {
    encrypt: process.env.MSSQL_ENCRYPT === "true",
    trustServerCertificate: process.env.MSSQL_TRUST_CERT !== "false",
    ...(process.env.MSSQL_INSTANCE ? { instanceName: process.env.MSSQL_INSTANCE } : {}),
  },
  pool: {
    max: 10,
    min: 0,
    idleTimeoutMillis: 30000,
  },
};

let _pool = null;

export async function getPool() {
  if (!_pool) {
    _pool = await sql.connect(config);
    console.log("[mssql] Connected to SQL Server:", config.server, "/", config.database);
  }
  return _pool;
}

export { sql };
