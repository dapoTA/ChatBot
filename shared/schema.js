import { mssqlTable, nvarchar, int, bit, datetime2, real } from "drizzle-orm/mssql-core";
import { createInsertSchema } from "drizzle-zod";

export const documents = mssqlTable("documents", {
  id: int("id").identity().primaryKey(),
  title: nvarchar("title", { length: "max" }).notNull(),
  content: nvarchar("content", { length: "max" }).notNull(),
  type: nvarchar("type", { length: "max" }).notNull(),
  url: nvarchar("url", { length: "max" }).notNull(),
  source: nvarchar("source", { length: "max" }).notNull().default("manual"),
  createdAt: datetime2("created_at").defaultNow(),
});

export const messages = mssqlTable("messages", {
  id: int("id").identity().primaryKey(),
  role: nvarchar("role", { length: "max" }).notNull(),
  content: nvarchar("content", { length: "max" }).notNull(),
  createdAt: datetime2("created_at").defaultNow(),
});

export const sharepointConfigs = mssqlTable("sharepoint_configs", {
  id: int("id").identity().primaryKey(),
  siteUrl: nvarchar("site_url", { length: "max" }).notNull(),
  domain: nvarchar("domain", { length: "max" }).notNull(),
  username: nvarchar("username", { length: "max" }).notNull(),
  password: nvarchar("password", { length: "max" }).notNull(),
  libraryName: nvarchar("library_name", { length: "max" }).notNull().default("Documents"),
  allowSelfSigned: bit("allow_self_signed").notNull().default(true),
  lastSyncedAt: datetime2("last_synced_at"),
  updatedAt: datetime2("updated_at").defaultNow(),
});

export const appSettings = mssqlTable("app_settings", {
  id: int("id").identity().primaryKey(),
  assistantName: nvarchar("assistant_name", { length: "max" }).notNull().default("ON-PNT® Assistant"),
  welcomeMessage: nvarchar("welcome_message", { length: "max" }).notNull().default("Ask me anything about your SharePoint documents."),
  notFoundMessage: nvarchar("not_found_message", { length: "max" }).notNull().default("I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator."),
  customInstructions: nvarchar("custom_instructions", { length: "max" }),
  temperature: real("temperature").notNull().default(0),
  topP: real("top_p").notNull().default(1),
  maxTokens: int("max_tokens").notNull().default(1500),
  frequencyPenalty: real("frequency_penalty").notNull().default(0),
  presencePenalty: real("presence_penalty").notNull().default(0),
  updatedAt: datetime2("updated_at").defaultNow(),
});

export const insertAppSettingsSchema = createInsertSchema(appSettings).omit({ id: true, updatedAt: true });
export const insertDocumentSchema = createInsertSchema(documents).omit({ id: true, createdAt: true });
export const insertMessageSchema = createInsertSchema(messages).omit({ id: true, createdAt: true });
export const insertSharepointConfigSchema = createInsertSchema(sharepointConfigs).omit({
  id: true,
  lastSyncedAt: true,
  updatedAt: true,
});
export const upsertSharepointConfigSchema = insertSharepointConfigSchema;
