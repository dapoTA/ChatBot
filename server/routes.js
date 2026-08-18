import { createServer } from "http";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { storage } from "./storage.js";
import { api } from "../shared/routes.js";
import { z } from "zod";
import { insertAppSettingsSchema } from "../shared/schema.js";
import { fetchLibraryItems, fetchDocumentContent, testSharepointConnection } from "./sharepoint.js";

const __dirname = typeof __filename !== "undefined"
  ? path.dirname(__filename)
  : path.dirname(fileURLToPath(import.meta.url));

//import OpenAI from "openai";

console.log("DEPLOYMENT:", process.env.AZURE_OPENAI_DEPLOYMENT);
console.log("API_VERSION:", process.env.AZURE_OPENAI_API_VERSION);
console.log("ENDPOINT:", process.env.AZURE_OPENAI_ENDPOINT);

//const openai = new OpenAI({
 // apiKey: process.env.AI_INTEGRATIONS_OPENAI_API_KEY,
 // baseURL: process.env.AI_INTEGRATIONS_OPENAI_BASE_URL,
//});

//import { AzureOpenAI } from "openai";

//const openai = new AzureOpenAI({
 // apiKey: process.env.AZURE_OPENAI_API_KEY,
 // endpoint: process.env.AZURE_OPENAI_ENDPOINT,      // https://techassurancetest01.openai.azure.com
 // apiVersion: process.env.AZURE_OPENAI_API_VERSION, // e.g. "2024-12-01-preview"
//});

// ─── Shared style constants ──────────────────────────────────────────────────
const STYLE_COLORS = [
  'red','blue','green','orange','purple','black','white','navy','maroon','teal',
  'darkred','darkblue','brown','pink','gold','gray','grey','silver','crimson',
  'coral','salmon','indigo','violet','cyan','magenta','lime','olive','turquoise',
];
const STYLE_SIZES = {
  'very large':  'font-size:1.6em;',
  'extra large': 'font-size:1.6em;',
  'large text':  'font-size:1.3em;',
  'large':       'font-size:1.3em;',
  'small text':  'font-size:0.85em;',
  'small':       'font-size:0.85em;',
};
// Keywords for matching natural language style descriptors
const _styleKeywords = `bold|italic|underline|very large|extra large|large|small|${STYLE_COLORS.join('|')}`;

// Parses raw custom instructions and returns a global response style object
// for patterns like "Always respond in bold green".
// Returns { color, bold, italic, underline, fontSize } or null.
export function extractGlobalResponseStyle(instructions) {
  if (!instructions) return null;
  const re = new RegExp(
    `\\b(?:respond|answer|write|reply|display|show)\\s+(?:all\\s+(?:answers?|responses?)\\s+)?in\\s+((?:(?:${_styleKeywords})\\s*)+(?:text)?)`,
    'gi'
  );
  let match;
  while ((match = re.exec(instructions)) !== null) {
    const desc = match[1].toLowerCase();
    const bold      = desc.includes('bold');
    const italic    = desc.includes('italic');
    const underline = desc.includes('underline');
    let color = null;
    for (const c of STYLE_COLORS) { if (desc.includes(c)) { color = c; break; } }
    let fontSize = null;
    for (const [label, css] of Object.entries(STYLE_SIZES)) { if (desc.includes(label)) { fontSize = css; break; } }
    if (bold || italic || underline || color || fontSize) {
      return { color, bold, italic, underline, fontSize };
    }
  }
  return null;
}

// Pre-processes custom instructions: detects natural-language style descriptions
// and rewrites them as explicit HTML the AI copies verbatim.
// Patterns that indicate the admin wrote a not-found fallback in their custom
// instructions. We strip these blocks before sending to the AI so the dedicated
// Not Found Message field is the single source of truth.
const NOT_FOUND_TRIGGERS = [
  /if\s+the\s+answer\s+isn['']t/i,
  /if\s+the\s+answer\s+is\s+not/i,
  /if\s+(?:you\s+)?cannot\s+find/i,
  /if\s+(?:you\s+)?can['']t\s+find/i,
  /if\s+(?:the\s+)?information\s+isn['']t/i,
  /if\s+(?:the\s+)?information\s+is\s+not/i,
  /if\s+it(?:'s|\s+is)\s+not\s+in/i,
  /if\s+there\s+is\s+no\s+(?:answer|information|relevant)/i,
  /if\s+(?:no|the)\s+(?:answer|information)[\s\S]{0,20}not\s+(?:found|available)/i,
  /answer\s+isn['']t\s+in\s+the\s+doc/i,
  /not\s+found.*respond\s+with/i,
];

function stripNotFoundFallback(instructions) {
  if (!instructions) return instructions;

  const lines = instructions.split('\n');
  const result = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];
    const isNotFoundLine = NOT_FOUND_TRIGGERS.some(p => p.test(line));

    if (isNotFoundLine) {
      // Skip this trigger line plus any immediately following lines that look
      // like the continuation of the fallback block (blank lines, quoted strings).
      i++;
      while (i < lines.length) {
        const next = lines[i].trim();
        const isContinuation =
          next === '' ||
          next.startsWith('"') ||
          next.startsWith("'") ||
          NOT_FOUND_TRIGGERS.some(p => p.test(next));
        if (!isContinuation) break;
        i++;
      }
    } else {
      result.push(line);
      i++;
    }
  }

  // Collapse any runs of 3+ blank lines left after stripping
  return result.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

function preprocessInstructions(instructions) {
  if (!instructions) return instructions;

  const COLORS = STYLE_COLORS;
  const colorList = COLORS.join('|');
  const SIZES = STYLE_SIZES;

  // Build a combined style string and HTML tag from a plain-language style descriptor
  function buildHtml(text, styleDesc) {
    const lower = styleDesc.toLowerCase();
    const isBold      = lower.includes('bold');
    const isItalic    = lower.includes('italic');
    const isUnderline = lower.includes('underline');

    let color = null;
    for (const c of COLORS) {
      if (lower.includes(c)) { color = c; break; }
    }

    let fontSize = null;
    for (const [label, css] of Object.entries(SIZES)) {
      if (lower.includes(label)) { fontSize = css; break; }
    }

    // Build inline style
    const styles = [];
    if (color)    styles.push(`color:${color};`);
    if (fontSize) styles.push(fontSize);
    if (isUnderline) styles.push('text-decoration:underline;');
    const styleAttr = styles.length ? ` style="${styles.join('')}"` : '';

    // Wrap with appropriate tags
    let html = text;
    if (styles.length || isItalic || isBold) {
      if (styles.length) html = `<span${styleAttr}>${html}</span>`;
      if (isItalic)      html = `<em>${html}</em>`;
      if (isBold)        html = `<strong>${html}</strong>`;
    } else {
      return null; // no recognisable style — leave untouched
    }
    return html;
  }

  const styleKeywords = `bold|italic|underline|large|small|very large|extra large|${colorList}`;

  // Pattern 1: "quoted text" in [style descriptors] [text?]
  // e.g. "Yes, I am eager..." in bold red text
  const quotedRe = new RegExp(
    `(?:["\u201C])([^"\u201D]+)(?:["\u201D])\\s+in\\s+((?:(?:${styleKeywords})\\s*)+(?:text)?)`,
    'gi'
  );
  instructions = instructions.replace(quotedRe, (match, quotedText, styleDesc) => {
    const html = buildHtml(quotedText, styleDesc);
    if (!html) return match;
    return `exactly this HTML (output it verbatim — never use markdown): ${html}`;
  });

  // Note: "respond in [style]" patterns are handled server-side via extractGlobalResponseStyle
  // and applied as CSS on the frontend — no AI instruction needed for those.

  return instructions;
}

export async function registerRoutes(httpServer, app) {

  // ─── CORS — allow SharePoint and any cross-origin consumer to load the widget ─
  app.use((req, res, next) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");
    if (req.method === "OPTIONS") return res.status(204).end();
    next();
  });

  // ─── /embed.js — serve chatbot-widget.js with server URL auto-injected ───────
  app.get("/embed.js", (req, res) => {
    // PUBLIC_URL env var overrides auto-detection — required when IIS terminates SSL
    // (iisnode uses a named pipe internally so req.protocol is always "http")
    const proto = req.headers["x-forwarded-proto"] || req.protocol || "http";
    const host = req.headers["x-forwarded-host"] || req.headers.host || "localhost";
    const serverUrl = process.env.PUBLIC_URL || `${proto}://${host}`;
    const widgetPath = path.join(__dirname, "../chatbot-widget.js");
    const script = fs.readFileSync(widgetPath, "utf8")
      .replace('"__CHATBOT_URL__"', JSON.stringify(serverUrl));
    res.setHeader("Content-Type", "application/javascript");
    res.setHeader("Cache-Control", "no-cache");
    res.send(script);
  });

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
      const { message, username, sessionId } = api.chat.send.input.parse(req.body);

      await storage.createMessage({ role: 'user', content: message });

      const docs = await storage.getDocuments();
      const appCfg = await storage.getAppSettings();
      const notFoundMessage = appCfg?.notFoundMessage
        ?? "I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator.";
      const rawInstructions = appCfg?.customInstructions?.trim() || null;
      // Strip any not-found fallback the admin may have written in their custom
      // instructions — the Not Found Message field is the sole source of truth.
      const strippedInstructions = stripNotFoundFallback(rawInstructions);
      const customInstructions = preprocessInstructions(strippedInstructions);

      const context = docs.map(d =>
        `[SOURCE]\nTitle: ${d.title}\nURL: ${d.url}\nType: ${d.type}\nContent: ${d.content}`
      ).join('\n\n---\n\n');

      const systemPrompt = `You are a SharePoint document assistant for this organization.
${customInstructions ? `\nOwner instructions — these are pre-approved by the organisation and govern your tone, style, format, prefix text, and response structure. Follow them precisely for every response:\n${customInstructions}\n` : ''}
NOT FOUND OVERRIDE: If the answer is not found in the approved documents below, you MUST respond with EXACTLY the following message — no more, no less. This overrides any alternative not-found or fallback wording that may appear in the owner instructions above:
${notFoundMessage}

For all other responses, the owner instructions above control your tone, style, and format. The restriction below applies only to factual content:

You may ONLY use factual content from the approved documents below to answer questions. Do not draw on your training data, general knowledge, or assumptions for factual answers. You MAY follow the owner instructions above for formatting, prefixes, tone, and response structure — those are pre-approved and not subject to this restriction.

Rules:
- If you cannot find a direct answer in the documents, use the exact not-found message above. Do not guess or infer.
- Keep answers clear and concise. Use bullet points where helpful.
- Every response that uses document content MUST end with a sources section in this exact format:

---
**Sources:**
- [Document Title](URL)

List every document you drew from. Use the exact title and URL from the source. This section is mandatory — never omit it.

Available documents:
${context || "No documents have been loaded yet. Please sync your SharePoint library in the Settings page."}`;

    //const completion = await openai.chat.completions.create({
     // model: process.env.AZURE_OPENAI_DEPLOYMENT,
      //messages: [
      //  { role: "system", content: systemPrompt },
      //  { role: "user", content: message },
      //],
    //});

                  const endpoint = new URL(
                `/openai/deployments/${process.env.AZURE_OPENAI_DEPLOYMENT}/chat/completions`,
                process.env.AZURE_OPENAI_ENDPOINT
              );

              endpoint.searchParams.set("api-version", process.env.AZURE_OPENAI_API_VERSION);

              console.log("Azure endpoint:", endpoint.toString());

              const response = await fetch(endpoint.toString(), {
                method: "POST",
                headers: {
                  "Content-Type": "application/json",
                  "api-key": process.env.AZURE_OPENAI_API_KEY,
                },
                body: JSON.stringify({
                  messages: [
                    { role: "system", content: systemPrompt },
                    { role: "user", content: message },
                  ],
                  temperature: appCfg?.temperature ?? 0,
                  top_p: appCfg?.topP ?? 1,
                  max_tokens: appCfg?.maxTokens ?? 1500,
                  frequency_penalty: appCfg?.frequencyPenalty ?? 0,
                  presence_penalty: appCfg?.presencePenalty ?? 0,
                }),
              });

              const raw = await response.text();
              console.log("Azure status:", response.status);
              console.log("Azure raw response:", raw);

              if (!response.ok) {
                throw new Error(`Azure OpenAI error: ${response.status} ${raw}`);
              }

              const completion = JSON.parse(raw);

      const aiResponse = completion.choices[0].message.content || "I couldn't generate a response.";
      const savedMessage = await storage.createMessage({ role: 'assistant', content: aiResponse });

      // Resolve username (priority order):
      //  1. ?spuser= param — injected by SharePoint embed scripts (on-prem & SPFx)
      //  2. iisnode's promoted AUTH_USER — populated by IIS after Windows
      //     authentication completes and available on every request
      const resolvedUsername =
        username ||
        req.headers['x-iisnode-auth_user'] ||
        null;

      // Write chat log (fire-and-forget, never blocks the response)
      if (appCfg?.enableChatLog) {
        storage.createChatLog({
          sessionId: sessionId || 'unknown',
          username:  resolvedUsername,
          userMessage:       message,
          assistantResponse: aiResponse,
        }).catch((err) => console.error('[chat-log] write failed:', err.message));
      }

      res.json(savedMessage);

    } catch (error) {
      console.error('Chat error:', error);
      res.status(500).json({ message: 'Failed to process chat message' });
    }
  });

  // ─── App settings routes (public — no credentials exposed) ──────────────────

  app.get("/api/settings", async (req, res) => {
    const settings = await storage.getAppSettings();
    const customInstructions = settings?.customInstructions ?? "";
    res.json({
      assistantName: settings?.assistantName ?? "ON-PNT® Assistant",
      welcomeMessage: settings?.welcomeMessage ?? "Ask me anything about your SharePoint documents.",
      notFoundMessage: settings?.notFoundMessage ?? "I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator.",
      customInstructions,
      responseStyle: extractGlobalResponseStyle(customInstructions),
      temperature: settings?.temperature ?? 0,
      topP: settings?.topP ?? 1,
      maxTokens: settings?.maxTokens ?? 1500,
      frequencyPenalty: settings?.frequencyPenalty ?? 0,
      presencePenalty: settings?.presencePenalty ?? 0,
      enableChatLog: settings?.enableChatLog ?? false,
    });
  });

  // ─── Chat log export ──────────────────────────────────────────────────────

  app.get("/api/chat-logs", async (req, res) => {
    try {
      const { from, to } = req.query;
      const logs = await storage.getChatLogs(from || null, to || null);

      const escape = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
      const header = 'id,created_at,username,session_id,user_message,assistant_response';
      const rows = logs.map((log) => [
        log.id,
        log.createdAt instanceof Date ? log.createdAt.toISOString() : (log.createdAt ?? ''),
        escape(log.username ?? ''),
        escape(log.sessionId),
        escape(log.userMessage),
        escape(log.assistantResponse),
      ].join(','));

      const csv = [header, ...rows].join('\r\n');
      const filename = `chat-logs-${new Date().toISOString().slice(0, 10)}.csv`;

      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      // BOM so Excel opens it correctly without a UTF-8 import wizard
      res.send('\uFEFF' + csv);
    } catch (err) {
      console.error('GET /api/chat-logs error:', err);
      res.status(500).json({ message: 'Failed to export chat logs.' });
    }
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
      console.error("POST /api/settings error:", err);
      return res.status(500).json({ message: "Failed to save settings." });
    }
  });

  // ─── SharePoint config routes ─────────────────────────────────────────────

  app.get(api.sharepoint.getConfig.path, async (req, res) => {
    const config = await storage.getSharepointConfig();
    const envControlled = {
      tenantId:     !!process.env.SHAREPOINT_TENANT_ID,
      clientId:     !!process.env.SHAREPOINT_CLIENT_ID,
      clientSecret: !!process.env.SHAREPOINT_CLIENT_SECRET,
    };
    if (config) {
      res.json({
        ...config,
        password: config.password ? "••••••••" : "",
        clientSecret: config.clientSecret ? "••••••••" : "",
        envControlled,
      });
    } else {
      res.json({ envControlled });
    }
  });

  app.post(api.sharepoint.saveConfig.path, async (req, res) => {
    try {
      const input = api.sharepoint.saveConfig.input.parse(req.body);

      const existing = await storage.getSharepointConfig();

      if (input.password === "••••••••") {
        if (existing) input.password = existing.password;
      }

      if (input.clientSecret === "••••••••") {
        if (existing) input.clientSecret = existing.clientSecret;
      }

      const config = await storage.upsertSharepointConfig(input);
      res.json({
        ...config,
        password: config.password ? "••••••••" : "",
        clientSecret: config.clientSecret ? "••••••••" : "",
      });
    } catch (err) {
      if (err instanceof z.ZodError) {
        return res.status(400).json({ message: err.errors[0].message, field: err.errors[0].path.join('.') });
      }
      throw err;
    }
  });

  app.post(api.sharepoint.testConnection.path, async (req, res) => {
    const dbConfig = await storage.getSharepointConfig();

    // If the frontend sent current form values, use those (so Test Connection
    // reflects the active tab without requiring a save first).
    let config;
    if (req.body && req.body.mode) {
      config = { ...dbConfig, ...req.body };
      // Restore real secrets if the masked placeholder was sent
      if (req.body.password === "••••••••" && dbConfig) config.password = dbConfig.password;
      if (req.body.clientSecret === "••••••••" && dbConfig) config.clientSecret = dbConfig.clientSecret;
      // Env-var credentials always win
      if (process.env.SHAREPOINT_TENANT_ID)     config.tenantId     = process.env.SHAREPOINT_TENANT_ID;
      if (process.env.SHAREPOINT_CLIENT_ID)     config.clientId     = process.env.SHAREPOINT_CLIENT_ID;
      if (process.env.SHAREPOINT_CLIENT_SECRET) config.clientSecret = process.env.SHAREPOINT_CLIENT_SECRET;
    } else {
      config = dbConfig;
    }

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
