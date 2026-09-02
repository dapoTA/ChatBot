import { pgTable, text, serial, boolean, timestamp, real, integer, uniqueIndex } from "drizzle-orm/pg-core";
import { sql } from "drizzle-orm";
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
  mode: text("mode").notNull().default("onprem"),
  siteUrl: text("site_url").notNull(),
  // Per-mode URL memory — each tab remembers its last URL independently
  siteUrlOnprem: text("site_url_onprem"),
  siteUrlOnline: text("site_url_online"),
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
  enableChatLog: boolean("enable_chat_log").notNull().default(false),
  assistantIcon: text("assistant_icon").notNull().default("message-circle"),
  theme: text("theme").notNull().default("teal"),
  customThemeColor: text("custom_theme_color"),
  launcherLabel: text("launcher_label").notNull().default("Ask inSite"),
  launcherPosition: text("launcher_position").notNull().default("bottom-right"),
  launcherStyle: text("launcher_style").notNull().default("bubble"),
  updatedAt: timestamp("updated_at").defaultNow(),
});

export const knowledgeSources = pgTable("knowledge_sources", {
  id: serial("id").primaryKey(),
  name: text("name").notNull(),
  libraryName: text("library_name"),
  sharepointMode: text("sharepoint_mode").notNull().default("inherit"),
  description: text("description").notNull().default(""),
  instructions: text("instructions").notNull().default(""),
  smeTeam: text("sme_team").notNull().default(""),
  contactMethod: text("contact_method").notNull().default(""),
  contactDetails: text("contact_details").notNull().default(""),
  escalationMessage: text("escalation_message").notNull().default(""),
  enabled: boolean("enabled").notNull().default(true),
  isPortalWide: boolean("is_portal_wide").notNull().default(false),
  createdAt: timestamp("created_at").defaultNow(),
  updatedAt: timestamp("updated_at").defaultNow(),
}, (table) => ({
  onePortalWideSource: uniqueIndex("knowledge_sources_one_portal_wide")
    .on(table.isPortalWide)
    .where(sql`${table.isPortalWide} = true`),
}));

export const chatLogs = pgTable("chat_logs", {
  id: serial("id").primaryKey(),
  sessionId: text("session_id").notNull(),
  username: text("username"),
  userMessage: text("user_message").notNull(),
  assistantResponse: text("assistant_response").notNull(),
  createdAt: timestamp("created_at").defaultNow(),
});

// enableChatLog is optional so partial saves (appearance-only, AI-params-only)
// don't accidentally overwrite the toggle back to false.
export const insertAppSettingsSchema = createInsertSchema(appSettings)
  .omit({ id: true, updatedAt: true })
  .extend({
    enableChatLog: z.boolean().optional(),
    assistantIcon: z.enum(["message-circle", "sparkles", "shield-check"]).optional(),
    theme: z.enum(["teal", "ocean", "slate", "amber", "custom"]).optional(),
    customThemeColor: z.preprocess(
      (value) => typeof value === "string" && value.trim() === "" ? null : value,
      z.string().trim().regex(/^#[0-9a-fA-F]{6}$/, "Use a 6-digit hexadecimal color such as #188B6A.").nullable().optional(),
    ),
    launcherLabel: z.string().trim().max(40).optional(),
    launcherPosition: z.enum(["bottom-right", "bottom-left"]).optional(),
    launcherStyle: z.enum(["bubble", "pill"]).optional(),
  });
export const insertKnowledgeSourceSchema = createInsertSchema(knowledgeSources)
  .omit({ id: true, createdAt: true, updatedAt: true })
  .extend({
    name: z.string().trim().min(1, "Source name is required").max(255),
    libraryName: z.string().trim().max(255).optional().nullable(),
    sharepointMode: z.enum(["inherit", "onprem", "online"]).optional().default("online"),
    description: z.string().optional().default(""),
    instructions: z.string().optional().default(""),
    smeTeam: z.string().optional().default(""),
    contactMethod: z.string().optional().default(""),
    contactDetails: z.string().optional().default(""),
    escalationMessage: z.string().optional().default(""),
    enabled: z.boolean().optional().default(true),
    isPortalWide: z.boolean().optional().default(false),
  });
export const updateKnowledgeSourceSchema = insertKnowledgeSourceSchema.partial();
export const insertChatLogSchema = createInsertSchema(chatLogs).omit({ id: true, createdAt: true });
export const insertDocumentSchema = createInsertSchema(documents).omit({ id: true, createdAt: true });
export const insertMessageSchema = createInsertSchema(messages).omit({ id: true, createdAt: true });
export const insertSharepointConfigSchema = createInsertSchema(sharepointConfigs).omit({
  id: true,
  lastSyncedAt: true,
  updatedAt: true,
});
export const upsertSharepointConfigSchema = insertSharepointConfigSchema;
