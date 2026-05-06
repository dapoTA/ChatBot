// SQL Server migration note:
// Drizzle ORM does not have a runtime adapter for SQL Server.
// This branch uses the `mssql` npm package directly for all database queries.
//
// To create the database tables, run:
//   node scripts/create-tables.js
//
// Ensure the following environment variables are set before running:
//   MSSQL_HOST, MSSQL_PORT, MSSQL_DATABASE, MSSQL_USER, MSSQL_PASSWORD, MSSQL_TRUST_CERT
