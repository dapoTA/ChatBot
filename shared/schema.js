import { z } from "zod";

export const insertDocumentSchema = z.object({
  title: z.string(),
  content: z.string(),
  type: z.string(),
  url: z.string(),
  source: z.string().default("manual"),
});

export const insertMessageSchema = z.object({
  role: z.string(),
  content: z.string(),
});

export const insertSharepointConfigSchema = z.object({
  siteUrl: z.string(),
  domain: z.string(),
  username: z.string(),
  password: z.string(),
  libraryName: z.string().default("Documents"),
  allowSelfSigned: z.boolean().default(true),
});

export const upsertSharepointConfigSchema = insertSharepointConfigSchema;

export const insertAppSettingsSchema = z.object({
  assistantName: z.string().default("ON-PNT\u00ae Assistant"),
  welcomeMessage: z.string().default("Ask me anything about your SharePoint documents."),
  notFoundMessage: z.string().default("I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator."),
  customInstructions: z.string().optional().nullable(),
  temperature: z.number().default(0),
  topP: z.number().default(1),
  maxTokens: z.number().int().default(1500),
  frequencyPenalty: z.number().default(0),
  presencePenalty: z.number().default(0),
});
