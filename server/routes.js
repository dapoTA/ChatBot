import { createServer } from "http";
import { storage } from "./storage.js";
import { api } from "@shared/routes";
import { z } from "zod";
import { insertAppSettingsSchema } from "@shared/schema";
import { fetchLibraryItems, fetchDocumentContent, testSharepointConnection } from "./sharepoint.js";
import OpenAI from "openai";

const openai = new OpenAI({
  apiKey: process.env.AI_INTEGRATIONS_OPENAI_API_KEY,
  baseURL: process.env.AI_INTEGRATIONS_OPENAI_BASE_URL,
});

export async function registerRoutes(httpServer, app) {

  // ─── Document routes ────────────────────────────────────────────────────────

  app.get(api.documents.list.path, async (req, res) => {
    const docs = await storage.getDocuments();
    res.json(docs);
  });

  app.post(api.documents.create.path, async (req, res) => {
    try {
      const input = api.documents.create.input.parse(req.body);
      const doc = await storage.createDocument(input);
      res.status(201).json(doc);
    } catch (err) {
      if (err instanceof z.ZodError) {
        return res.status(400).json({ message: err.errors[0].message, field: err.errors[0].path.join('.') });
      }
      throw err;
    }
  });

  app.delete(api.documents.delete.path, async (req, res) => {
    const id = parseInt(req.params.id);
    await storage.deleteDocument(id);
    res.status(204).send();
  });

  // ─── Chat routes ─────────────────────────────────────────────────────────────

  app.get(api.chat.history.path, async (req, res) => {
    const messages = await storage.getMessages();
    res.json(messages);
  });

  app.delete(api.chat.clear.path, async (req, res) => {
    await storage.clearMessages();
    res.status(204).send();
  });

  app.post(api.chat.send.path, async (req, res) => {
    try {
      const { message } = api.chat.send.input.parse(req.body);

      await storage.createMessage({ role: 'user', content: message });

      const docs = await storage.getDocuments();
      const appCfg = await storage.getAppSettings();
      const notFoundMessage = appCfg?.notFoundMessage
        ?? "I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator.";

      const context = docs.map(d =>
        `[SOURCE]\nTitle: ${d.title}\nURL: ${d.url}\nType: ${d.type}\nContent: ${d.content}`
      ).join('\n\n---\n\n');

      const systemPrompt = `You are a helpful SharePoint assistant for this organization.

Your job is to answer questions based ONLY on the documents provided below.
If the answer is not found in the documents, respond with EXACTLY this message (do not modify it):
"${notFoundMessage}"

Rules:
- Always cite the specific document(s) you used to answer the question.
- When referencing a document, include a Markdown link using the exact URL provided in the source: [Document Title](URL)
- If multiple documents are relevant, cite all of them.
- Keep answers clear and concise. Use bullet points where helpful.
- Never make up information not present in the documents.

Available documents:
${context || "No documents have been loaded yet. Please sync your SharePoint library in the Settings page."}`;

      const completion = await openai.chat.completions.create({
        model: "gpt-5.1",
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: message },
        ],
      });

      const aiResponse = completion.choices[0].message.content || "I couldn't generate a response.";
      const savedMessage = await storage.createMessage({ role: 'assistant', content: aiResponse });
      res.json(savedMessage);

    } catch (error) {
      console.error('Chat error:', error);
      res.status(500).json({ message: 'Failed to process chat message' });
    }
  });

  // ─── App settings routes (public — no credentials exposed) ──────────────────

  app.get("/api/settings", async (req, res) => {
    const settings = await storage.getAppSettings();
    res.json({
      assistantName: settings?.assistantName ?? "ON-PNT® Assistant",
      welcomeMessage: settings?.welcomeMessage ?? "Ask me anything about your SharePoint documents.",
      notFoundMessage: settings?.notFoundMessage ?? "I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator.",
    });
  });

  app.post("/api/settings", async (req, res) => {
    try {
      const input = insertAppSettingsSchema.parse(req.body);
      const saved = await storage.upsertAppSettings(input);
      res.json(saved);
    } catch (err) {
      if (err instanceof z.ZodError) {
        return res.status(400).json({ message: err.errors[0].message });
      }
      throw err;
    }
  });

  // ─── SharePoint config routes ─────────────────────────────────────────────

  app.get(api.sharepoint.getConfig.path, async (req, res) => {
    const config = await storage.getSharepointConfig();
    if (config) {
      res.json({ ...config, password: config.password ? "••••••••" : "" });
    } else {
      res.json(null);
    }
  });

  app.post(api.sharepoint.saveConfig.path, async (req, res) => {
    try {
      const input = api.sharepoint.saveConfig.input.parse(req.body);

      if (input.password === "••••••••") {
        const existing = await storage.getSharepointConfig();
        if (existing) {
          input.password = existing.password;
        }
      }

      const config = await storage.upsertSharepointConfig(input);
      res.json({ ...config, password: "••••••••" });
    } catch (err) {
      if (err instanceof z.ZodError) {
        return res.status(400).json({ message: err.errors[0].message, field: err.errors[0].path.join('.') });
      }
      throw err;
    }
  });

  app.post(api.sharepoint.testConnection.path, async (req, res) => {
    const config = await storage.getSharepointConfig();
    if (!config) {
      return res.json({ success: false, message: "No SharePoint configuration saved yet." });
    }
    const result = await testSharepointConnection(config);
    res.json(result);
  });

  app.post(api.sharepoint.sync.path, async (req, res) => {
    const config = await storage.getSharepointConfig();
    if (!config) {
      return res.status(400).json({ message: "No SharePoint configuration found. Please save your settings first." });
    }

    let synced = 0;
    let failed = 0;

    try {
      await storage.deleteDocumentsBySource("sharepoint");

      const items = await fetchLibraryItems(config);
      console.log(`SharePoint sync: found ${items.length} files in "${config.libraryName}"`);

      for (const item of items) {
        try {
          const docWithContent = await fetchDocumentContent(config, item);
          await storage.createDocument({
            title: docWithContent.title,
            content: docWithContent.content,
            type: "document",
            url: docWithContent.url,
            source: "sharepoint",
          });
          synced++;
        } catch (err) {
          console.error(`Failed to sync "${item.title}":`, err?.message);
          failed++;
        }
      }

      await storage.updateSharepointSyncTime();

      res.json({
        synced,
        failed,
        message: `Sync complete. ${synced} document${synced !== 1 ? 's' : ''} imported${failed > 0 ? `, ${failed} failed` : ''}.`,
      });

    } catch (error) {
      console.error("SharePoint sync error:", error);
      res.status(500).json({ message: error?.message || "Sync failed. Check your SharePoint configuration." });
    }
  });

  // ─── Seed demo data (only if no documents at all) ────────────────────────

  async function seedDatabase() {
    const docs = await storage.getDocuments();
    if (docs.length === 0) {
      await storage.createDocument({
        title: "IT Support Policy",
        content: "All employees must change their passwords every 90 days. For technical support, create a ticket at support.company.com or call extension 5555. Standard response time is 24 hours.",
        type: "document",
        url: "http://sharepoint.company.com/sites/it/policies/support.docx",
        source: "manual",
      });
      await storage.createDocument({
        title: "Q4 Marketing Plan",
        content: "The Q4 marketing strategy focuses on social media engagement and email campaigns. Key dates: Nov 1st - Holiday Campaign Launch, Dec 15th - Year End Review. Budget allocated: $50,000.",
        type: "document",
        url: "http://sharepoint.company.com/sites/marketing/2024/q4-plan.pptx",
        source: "manual",
      });
      await storage.createDocument({
        title: "Project Alpha Tasks",
        content: "1. UI Design Phase - Status: In Progress, Owner: Sarah\n2. Database Migration - Status: Pending, Owner: Mike\n3. User Testing - Status: Not Started",
        type: "list-item",
        url: "http://sharepoint.company.com/sites/projects/lists/tasks/123",
        source: "manual",
      });
      console.log('Seeded database with example SharePoint content');
    }
  }

  seedDatabase();

  return httpServer;
}
