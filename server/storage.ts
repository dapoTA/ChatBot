import { db } from "./db";
import { documents, messages, type InsertDocument, type InsertMessage } from "@shared/schema";
import { eq, desc } from "drizzle-orm";

export interface IStorage {
  getDocuments(): Promise<(typeof documents.$inferSelect)[]>;
  createDocument(doc: InsertDocument): Promise<typeof documents.$inferSelect>;
  deleteDocument(id: number): Promise<void>;
  
  getMessages(): Promise<(typeof messages.$inferSelect)[]>;
  createMessage(msg: InsertMessage): Promise<typeof messages.$inferSelect>;
  clearMessages(): Promise<void>;
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
}

export const storage = new DatabaseStorage();
