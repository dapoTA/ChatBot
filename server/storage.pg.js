import { db } from "./db.js";
import {
  documents,
  messages,
  sharepointConfigs,
  appSettings,
  knowledgeSources,
  chatLogs,
} from "../shared/schema.js";
import { eq, desc, gte, lte, and } from "drizzle-orm";
import { encrypt, decrypt } from "./crypto.js";

// Fields in sharepoint_configs that are encrypted at rest
const ENCRYPTED_FIELDS = ["password", "clientSecret"];

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
    const row = configs[0] ?? null;
    if (!row) return null;
    // Decrypt sensitive fields before returning to the application layer
    const decrypted = { ...row };
    for (const field of ENCRYPTED_FIELDS) {
      if (decrypted[field]) decrypted[field] = decrypt(decrypted[field]);
    }
    return decrypted;
  }

  async upsertSharepointConfig(config) {
    // Encrypt sensitive fields before writing to the database
    const toStore = { ...config };
    for (const field of ENCRYPTED_FIELDS) {
      if (toStore[field] && toStore[field] !== "••••••••") {
        toStore[field] = encrypt(toStore[field]);
      }
    }

    const existing = await db.select().from(sharepointConfigs).limit(1);
    const existingRow = existing[0] ?? null;

    if (existingRow) {
      const [updated] = await db
        .update(sharepointConfigs)
        .set({ ...toStore, updatedAt: new Date() })
        .where(eq(sharepointConfigs.id, existingRow.id))
        .returning();
      // Return with decrypted fields so callers get plain-text values
      const out = { ...updated };
      for (const field of ENCRYPTED_FIELDS) {
        if (out[field]) out[field] = decrypt(out[field]);
      }
      return out;
    } else {
      const [created] = await db.insert(sharepointConfigs).values(toStore).returning();
      const out = { ...created };
      for (const field of ENCRYPTED_FIELDS) {
        if (out[field]) out[field] = decrypt(out[field]);
      }
      return out;
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

  async getKnowledgeSources() {
    return await db
      .select()
      .from(knowledgeSources)
      .orderBy(desc(knowledgeSources.isPortalWide), knowledgeSources.name, knowledgeSources.id);
  }

  async createKnowledgeSource(source) {
    const [created] = await db.insert(knowledgeSources).values(source).returning();
    return created;
  }

  async updateKnowledgeSource(id, source) {
    const [updated] = await db
      .update(knowledgeSources)
      .set({ ...source, updatedAt: new Date() })
      .where(eq(knowledgeSources.id, id))
      .returning();
    return updated ?? null;
  }

  async deleteKnowledgeSource(id) {
    await db.delete(knowledgeSources).where(eq(knowledgeSources.id, id));
  }

  async createChatLog(log) {
    const [row] = await db.insert(chatLogs).values(log).returning();
    return row;
  }

  async getChatLogs(from, to) {
    const conditions = [];
    if (from) conditions.push(gte(chatLogs.createdAt, new Date(from)));
    if (to) {
      const toDate = new Date(to);
      toDate.setDate(toDate.getDate() + 1);
      conditions.push(lte(chatLogs.createdAt, toDate));
    }
    let q = db.select().from(chatLogs);
    if (conditions.length === 1) q = q.where(conditions[0]);
    else if (conditions.length > 1) q = q.where(and(...conditions));
    return await q.orderBy(desc(chatLogs.createdAt));
  }
}

export const storage = new DatabaseStorage();
