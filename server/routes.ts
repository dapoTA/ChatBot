import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import { api } from "@shared/routes";
import { z } from "zod";
import { registerChatRoutes } from "./replit_integrations/chat";
import OpenAI from "openai";

// Initialize OpenAI client using integration env vars
const openai = new OpenAI({
  apiKey: process.env.AI_INTEGRATIONS_OPENAI_API_KEY,
  baseURL: process.env.AI_INTEGRATIONS_OPENAI_BASE_URL,
});

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {
  // Document management routes
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
        return res.status(400).json({
          message: err.errors[0].message,
          field: err.errors[0].path.join('.'),
        });
      }
      throw err;
    }
  });

  app.delete(api.documents.delete.path, async (req, res) => {
    const id = parseInt(req.params.id);
    await storage.deleteDocument(id);
    res.status(204).send();
  });

  // Chat routes
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
      
      // Save user message
      await storage.createMessage({ role: 'user', content: message });

      // Get context from documents
      const docs = await storage.getDocuments();
      const context = docs.map(d => 
        `[${d.type.toUpperCase()}] ${d.title} (${d.url}):\n${d.content}`
      ).join('\n\n');

      // Construct prompt for RAG
      const systemPrompt = `You are a helpful SharePoint assistant. 
      Answer the user's question based ONLY on the following context.
      If the answer is not in the context, say so politely.
      
      Context:
      ${context}`;

      // Call OpenAI
      const completion = await openai.chat.completions.create({
        model: "gpt-5.1",
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: message }
        ],
      });

      const aiResponse = completion.choices[0].message.content || "I couldn't generate a response.";

      // Save and return AI message
      const savedMessage = await storage.createMessage({ role: 'assistant', content: aiResponse });
      res.json(savedMessage);

    } catch (error) {
      console.error('Chat error:', error);
      res.status(500).json({ message: 'Failed to process chat message' });
    }
  });

  // Seed data function
  async function seedDatabase() {
    const docs = await storage.getDocuments();
    if (docs.length === 0) {
      await storage.createDocument({
        title: "IT Support Policy",
        content: "All employees must change their passwords every 90 days. For technical support, create a ticket at support.company.com or call extension 5555. Standard response time is 24 hours.",
        type: "document",
        url: "/sites/it/policies/support.docx"
      });
      await storage.createDocument({
        title: "Q4 Marketing Plan",
        content: "The Q4 marketing strategy focuses on social media engagement and email campaigns. Key dates: Nov 1st - Holiday Campaign Launch, Dec 15th - Year End Review. Budget allocated: $50,000.",
        type: "document",
        url: "/sites/marketing/2024/q4-plan.pptx"
      });
      await storage.createDocument({
        title: "Project Alpha Tasks",
        content: "1. UI Design Phase - Status: In Progress, Owner: Sarah\n2. Database Migration - Status: Pending, Owner: Mike\n3. User Testing - Status: Not Started",
        type: "list-item",
        url: "/sites/projects/lists/tasks/123"
      });
      console.log('Seeded database with example SharePoint content');
    }
  }

  // Run seeder
  seedDatabase();

  return httpServer;
}
