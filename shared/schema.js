import { pgTable, text, serial, boolean, timestamp, real, integer } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";

export const documents = pgTable("documents", {
  id: serial("id").primaryKey(),
  title: text("title").notNull(),
  content: text("content").notNull(),
  type: text("type").notNull(),
  url: text("url").notNull(),
  source: text("source").notNull().default("manual"),
  createdAt: timestamp("created_at").defaultNow(),
});

export const messages = pgTable("messages", {
  id: serial("id").primaryKey(),
  role: text("role").notNull(),
  content: text("content").notNull(),
  createdAt: timestamp("created_at").defaultNow(),
});

export const sharepointConfigs = pgTable("sharepoint_configs", {
  id: serial("id").primaryKey(),
  mode: text("mode").notNull().default("onprem"),
  siteUrl: text("site_url").notNull(),
  libraryName: text("library_name").notNull().default("Documents"),
  // On-premises NTLM fields
  domain: text("domain").notNull().default(""),
  username: text("username").notNull().default(""),
  password: text("password").notNull().default(""),
  allowSelfSigned: boolean("allow_self_signed").notNull().default(true),
  // SharePoint Online OAuth fields
  tenantId: text("tenant_id"),
  clientId: text("client_id"),
  clientSecret: text("client_secret"),
  lastSyncedAt: timestamp("last_synced_at"),
  updatedAt: timestamp("updated_at").defaultNow(),
});

export const appSettings = pgTable("app_settings", {
  id: serial("id").primaryKey(),
  assistantName: text("assistant_name").notNull().default("ON-PNT® Assistant"),
  welcomeMessage: text("welcome_message").notNull().default("Ask me anything about your SharePoint documents."),
  notFoundMessage: text("not_found_message").notNull().default("I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator."),
  customInstructions: text("custom_instructions"),
  temperature: real("temperature").notNull().default(0),
  topP: real("top_p").notNull().default(1),
  maxTokens: integer("max_tokens").notNull().default(1500),
  frequencyPenalty: real("frequency_penalty").notNull().default(0),
  presencePenalty: real("presence_penalty").notNull().default(0),
  updatedAt: timestamp("updated_at").defaultNow(),
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
