import { db } from "./db.js";
import { documents, messages, sharepointConfigs, appSettings } from "../shared/schema.js";
import { eq, desc } from "drizzle-orm";

export class DatabaseStorage {
  async getDocuments() {
    return await db.select().from(documents).orderBy(desc(documents.createdAt));
  }

  async createDocument(doc) {
    const [document] = await db.insert(documents).values(doc).returning();
    return document;
  }

  async deleteDocument(id) {
    await db.delete(documents).where(eq(documents.id, id));
  }

  async deleteDocumentsBySource(source) {
    await db.delete(documents).where(eq(documents.source, source));
  }

  async getMessages() {
    return await db.select().from(messages).orderBy(messages.createdAt);
  }

  async createMessage(msg) {
    const [message] = await db.insert(messages).values(msg).returning();
    return message;
  }

  async clearMessages() {
    await db.delete(messages);
  }

  async getSharepointConfig() {
    const configs = await db.select().from(sharepointConfigs).limit(1);
    return configs[0] ?? null;
  }

  async upsertSharepointConfig(config) {
    const existing = await this.getSharepointConfig();
    if (existing) {
      const [updated] = await db
        .update(sharepointConfigs)
        .set({ ...config, updatedAt: new Date() })
        .where(eq(sharepointConfigs.id, existing.id))
        .returning();
      return updated;
    } else {
      const [created] = await db.insert(sharepointConfigs).values(config).returning();
      return created;
    }
  }

  async updateSharepointSyncTime() {
    const existing = await this.getSharepointConfig();
    if (existing) {
      await db
        .update(sharepointConfigs)
        .set({ lastSyncedAt: new Date() })
        .where(eq(sharepointConfigs.id, existing.id));
    }
  }

  async getAppSettings() {
    const rows = await db.select().from(appSettings).limit(1);
    return rows[0] ?? null;
  }

  async upsertAppSettings(settings) {
    const existing = await this.getAppSettings();
    if (existing) {
      const [updated] = await db
        .update(appSettings)
        .set({ ...settings, updatedAt: new Date() })
        .where(eq(appSettings.id, existing.id))
        .returning();
      return updated;
    } else {
      const [created] = await db.insert(appSettings).values(settings).returning();
      return created;
    }
  }
}

export const storage = new DatabaseStorage();
