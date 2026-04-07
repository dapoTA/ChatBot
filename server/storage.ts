import { db } from "./db";
import { documents, messages, sharepointConfigs, type InsertDocument, type InsertMessage, type InsertSharepointConfig } from "@shared/schema";
import { eq, desc } from "drizzle-orm";

export interface IStorage {
  getDocuments(): Promise<(typeof documents.$inferSelect)[]>;
  createDocument(doc: InsertDocument): Promise<typeof documents.$inferSelect>;
  deleteDocument(id: number): Promise<void>;
  deleteDocumentsBySource(source: string): Promise<void>;

  getMessages(): Promise<(typeof messages.$inferSelect)[]>;
  createMessage(msg: InsertMessage): Promise<typeof messages.$inferSelect>;
  clearMessages(): Promise<void>;

  getSharepointConfig(): Promise<typeof sharepointConfigs.$inferSelect | null>;
  upsertSharepointConfig(config: InsertSharepointConfig): Promise<typeof sharepointConfigs.$inferSelect>;
  updateSharepointSyncTime(): Promise<void>;
}

export class DatabaseStorage implements IStorage {
  async getDocuments() {
    return await db.select().from(documents).orderBy(desc(documents.createdAt));
  }

  async createDocument(doc: InsertDocument) {
    const [document] = await db.insert(documents).values(doc).returning();
    return document;
  }

  async deleteDocument(id: number) {
    await db.delete(documents).where(eq(documents.id, id));
  }

  async deleteDocumentsBySource(source: string) {
    await db.delete(documents).where(eq(documents.source, source));
  }

  async getMessages() {
    return await db.select().from(messages).orderBy(messages.createdAt);
  }

  async createMessage(msg: InsertMessage) {
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

  async upsertSharepointConfig(config: InsertSharepointConfig) {
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
}

export const storage = new DatabaseStorage();
