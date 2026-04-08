import { pgTable, text, serial, boolean, timestamp } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

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
  siteUrl: text("site_url").notNull(),
  domain: text("domain").notNull(),
  username: text("username").notNull(),
  password: text("password").notNull(),
  libraryName: text("library_name").notNull().default("Documents"),
  allowSelfSigned: boolean("allow_self_signed").notNull().default(true),
  lastSyncedAt: timestamp("last_synced_at"),
  updatedAt: timestamp("updated_at").defaultNow(),
});

export const appSettings = pgTable("app_settings", {
  id: serial("id").primaryKey(),
  assistantName: text("assistant_name").notNull().default("ON-PNT® Assistant"),
  welcomeMessage: text("welcome_message").notNull().default("Ask me anything about your SharePoint documents."),
  notFoundMessage: text("not_found_message").notNull().default("I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator."),
  updatedAt: timestamp("updated_at").defaultNow(),
});

export const insertAppSettingsSchema = createInsertSchema(appSettings).omit({ id: true, updatedAt: true });
export type AppSettings = typeof appSettings.$inferSelect;
export type InsertAppSettings = z.infer<typeof insertAppSettingsSchema>;

export const insertDocumentSchema = createInsertSchema(documents).omit({ id: true, createdAt: true });
export const insertMessageSchema = createInsertSchema(messages).omit({ id: true, createdAt: true });
export const insertSharepointConfigSchema = createInsertSchema(sharepointConfigs).omit({
  id: true,
  lastSyncedAt: true,
  updatedAt: true,
});
export const upsertSharepointConfigSchema = insertSharepointConfigSchema;

export type Document = typeof documents.$inferSelect;
export type InsertDocument = z.infer<typeof insertDocumentSchema>;
export type Message = typeof messages.$inferSelect;
export type InsertMessage = z.infer<typeof insertMessageSchema>;
export type SharepointConfig = typeof sharepointConfigs.$inferSelect;
export type InsertSharepointConfig = z.infer<typeof insertSharepointConfigSchema>;

export type CreateDocumentRequest = InsertDocument;
export type ChatRequest = { message: string };
