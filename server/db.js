import mssql from "mssql";

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
export { mssql as sql };
