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

// Converts assistant HTML markup into readable text for CSV exports.
// Stored responses remain unchanged because the widget still needs the markup.
export function htmlToPlainText(value) {
  if (value == null) return '';

  return String(value)
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\s*li\b[^>]*>/gi, '- ')
    .replace(/<\s*\/\s*(p|div|li|h[1-6]|tr)\s*>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/gi, "'")
    .replace(/&mdash;/gi, '—')
    .replace(/&ndash;/gi, '–')
    .replace(/&hellip;/gi, '…')
    .replace(/&bull;/gi, '•')
    .replace(/&ldquo;/gi, '“')
    .replace(/&rdquo;/gi, '”')
    .replace(/&lsquo;/gi, '‘')
    .replace(/&rsquo;/gi, '’')
    .replace(/&#(\d+);/g, (_, code) => {
      const point = Number(code);
      return point >= 0 && point <= 0x10ffff ? String.fromCodePoint(point) : _;
    })
    .replace(/&#x([0-9a-f]+);/gi, (_, code) => {
      const point = Number.parseInt(code, 16);
      return point >= 0 && point <= 0x10ffff ? String.fromCodePoint(point) : _;
    })
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
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

const DEFAULT_KNOWLEDGE_SOURCES = [
  {
    name: "All Portal Sources",
    libraryName: null,
    description: "Searches the full configured SharePoint site collection.",
    instructions: "",
    smeTeam: "",
    contactMethod: "",
    contactDetails: "",
    escalationMessage: "",
    enabled: true,
    isPortalWide: true,
  },
  {
    name: "PTO",
    libraryName: "PTO",
    description: "Searches the PTO and leave information library.",
    instructions: "",
    smeTeam: "",
    contactMethod: "",
    contactDetails: "",
    escalationMessage: "",
    enabled: true,
    isPortalWide: false,
  },
  {
    name: "HR",
    libraryName: "HR",
    description: "Searches the HR information library.",
    instructions: "",
    smeTeam: "",
    contactMethod: "",
    contactDetails: "",
    escalationMessage: "",
    enabled: true,
    isPortalWide: false,
  },
  {
    name: "Company Policies",
    libraryName: "Company Policies",
    description: "Searches the company policies library.",
    instructions: "",
    smeTeam: "",
    contactMethod: "",
    contactDetails: "",
    escalationMessage: "",
    enabled: true,
    isPortalWide: false,
  },
];

function normaliseSourceValue(value) {
  return String(value ?? "").trim().toLowerCase();
}

function knowledgeSourceDocumentKey(source) {
  return `knowledge-source:${source.id}`;
}

function requestUsername(req) {
  return String(
    req.headers["x-iisnode-auth_user"]
    || req.user?.username
    || "",
  ).trim();
}

function requireKnowledgeSourceAdmin(req, res, next) {
  if (process.env.NODE_ENV !== "production") return next();

  const username = requestUsername(req);
  if (!username) {
    return res.status(401).json({ message: "Windows authentication is required." });
  }

  const allowedAdmins = String(process.env.ADMIN_USERS || "")
    .split(/[,\n;]/)
    .map(normaliseSourceValue)
    .filter(Boolean);
  if (allowedAdmins.length === 0) {
    return res.status(503).json({
      message: "Knowledge source administration is not configured. Set ADMIN_USERS on the server.",
    });
  }

  const normalisedUsername = normaliseSourceValue(username);
  const shortUsername = normalisedUsername.split(/[\\/]/).pop();
  const permitted = allowedAdmins.some((admin) => {
    const shortAdmin = admin.split(/[\\/]/).pop();
    return admin === normalisedUsername || shortAdmin === shortUsername;
  });

  if (!permitted) {
    return res.status(403).json({ message: "Administrator access is required." });
  }
  next();
}

export function sourceEscalationText(source) {
  if (!source) return "";

  const routing = [];
  if (source.escalationMessage?.trim()) routing.push(source.escalationMessage.trim());

  const contact = [
    source.smeTeam?.trim(),
    source.contactMethod?.trim(),
    source.contactDetails?.trim(),
  ].filter(Boolean);

  if (contact.length) {
    routing.push(`Subject-matter expert contact: ${contact.join(" · ")}`);
  }

  return routing.join("\n\n").trim();
}

export function documentsForSource(documents, source, sharepointConfig, knowledgeSources = []) {
  if (!source) return documents;
  if (source.isPortalWide) {
    if (knowledgeSources.length === 0) return documents;
    const enabledSourceKeys = new Set(
      knowledgeSources
        .filter((item) => item.enabled && !item.isPortalWide)
        .map((item) => knowledgeSourceDocumentKey(item)),
    );
    return documents.filter((document) => {
      const documentSource = String(document.source ?? "");
      if (normaliseSourceValue(documentSource) === "sharepoint") {
        const configuredLibrary = normaliseSourceValue(sharepointConfig?.libraryName);
        const configuredSource = knowledgeSources.find(
          (item) =>
            !item.isPortalWide
            && normaliseSourceValue(item.libraryName) === configuredLibrary,
        );
        return configuredSource ? configuredSource.enabled : true;
      }
      return !documentSource.startsWith("knowledge-source:")
        || enabledSourceKeys.has(documentSource);
    });
  }

  const sourceNames = new Set(
    [source.name, source.libraryName, knowledgeSourceDocumentKey(source)]
      .map(normaliseSourceValue)
      .filter(Boolean),
  );
  const configuredLibrary = normaliseSourceValue(sharepointConfig?.libraryName);

  return documents.filter((document) => {
    const documentSource = normaliseSourceValue(document.source);
    if (sourceNames.has(documentSource)) return true;

    // Existing installations labelled synced documents as "sharepoint".
    // Treat those documents as belonging to the configured library so the
    // new source selector remains backward compatible.
    return documentSource === "sharepoint"
      && configuredLibrary
      && sourceNames.has(configuredLibrary);
  });
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

  // ─── Knowledge source administration ──────────────────────────────────────
  app.get(api.knowledgeSources.options.path, async (_req, res) => {
    const sources = await storage.getKnowledgeSources();
    res.json(
      sources
        .filter((source) => source.enabled)
        .map((source) => ({
          id: source.id,
          name: source.name,
          description: source.description ?? "",
          isPortalWide: source.isPortalWide,
        })),
    );
  });

  app.get(api.knowledgeSources.list.path, requireKnowledgeSourceAdmin, async (_req, res) => {
    res.json(await storage.getKnowledgeSources());
  });

  app.post(api.knowledgeSources.create.path, requireKnowledgeSourceAdmin, async (req, res) => {
    try {
      const input = api.knowledgeSources.create.input.parse(req.body);
      if (input.isPortalWide) {
        return res.status(400).json({
          message: "All Portal Sources is managed automatically and cannot be added.",
          field: "isPortalWide",
        });
      }
      if (!input.libraryName?.trim()) {
        return res.status(400).json({
          message: "A SharePoint library name is required for a named source.",
          field: "libraryName",
        });
      }

      const sources = await storage.getKnowledgeSources();
      const duplicate = sources.find(
        (source) => normaliseSourceValue(source.name) === normaliseSourceValue(input.name),
      );
      if (duplicate) {
        return res.status(400).json({
          message: "A knowledge source with that name already exists.",
          field: "name",
        });
      }

      const created = await storage.createKnowledgeSource({
        ...input,
        name: input.name.trim(),
        libraryName: input.libraryName.trim(),
      });
      res.status(201).json(created);
    } catch (err) {
      if (err instanceof z.ZodError) {
        return res.status(400).json({
          message: err.errors[0].message,
          field: err.errors[0].path.join("."),
        });
      }
      throw err;
    }
  });

  app.patch(api.knowledgeSources.update.path, requireKnowledgeSourceAdmin, async (req, res) => {
    try {
      const id = Number.parseInt(req.params.id, 10);
      const sources = await storage.getKnowledgeSources();
      const existing = sources.find((source) => source.id === id);
      if (!existing) {
        return res.status(404).json({ message: "Knowledge source not found." });
      }

      const input = api.knowledgeSources.update.input.parse(req.body);
      if (
        Object.prototype.hasOwnProperty.call(input, "isPortalWide")
        && input.isPortalWide !== existing.isPortalWide
      ) {
        return res.status(400).json({
          message: "The portal-wide source type cannot be changed.",
          field: "isPortalWide",
        });
      }
      const merged = {
        ...existing,
        ...input,
        name: (input.name ?? existing.name).trim(),
        libraryName: input.libraryName !== undefined
          ? input.libraryName?.trim() || null
          : existing.libraryName,
      };

      if (existing.isPortalWide) {
        merged.name = existing.name;
        merged.libraryName = null;
        merged.enabled = true;
        merged.isPortalWide = true;
      } else if (!merged.libraryName) {
        return res.status(400).json({
          message: "A SharePoint library name is required for a named source.",
          field: "libraryName",
        });
      }

      const duplicate = sources.find(
        (source) =>
          source.id !== id
          && normaliseSourceValue(source.name) === normaliseSourceValue(merged.name),
      );
      if (duplicate) {
        return res.status(400).json({
          message: "A knowledge source with that name already exists.",
          field: "name",
        });
      }

      const updated = await storage.updateKnowledgeSource(id, {
        name: merged.name,
        libraryName: merged.libraryName,
        description: merged.description ?? "",
        instructions: merged.instructions ?? "",
        smeTeam: merged.smeTeam ?? "",
        contactMethod: merged.contactMethod ?? "",
        contactDetails: merged.contactDetails ?? "",
        escalationMessage: merged.escalationMessage ?? "",
        enabled: merged.enabled !== false,
        isPortalWide: merged.isPortalWide === true,
      });
      res.json(updated);
    } catch (err) {
      if (err instanceof z.ZodError) {
        return res.status(400).json({
          message: err.errors[0].message,
          field: err.errors[0].path.join("."),
        });
      }
      throw err;
    }
  });

  app.delete(api.knowledgeSources.delete.path, requireKnowledgeSourceAdmin, async (req, res) => {
    const id = Number.parseInt(req.params.id, 10);
    const sources = await storage.getKnowledgeSources();
    const source = sources.find((item) => item.id === id);
    if (!source) return res.status(404).json({ message: "Knowledge source not found." });
    if (source.isPortalWide) {
      return res.status(400).json({
        message: "All Portal Sources is required and cannot be removed.",
      });
    }
    await storage.deleteDocumentsBySource(knowledgeSourceDocumentKey(source));
    await storage.deleteKnowledgeSource(id);
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
      const { message, username, sessionId, sourceId } = api.chat.send.input.parse(req.body);

      const allDocuments = await storage.getDocuments();
      const knowledgeSources = await storage.getKnowledgeSources();
      let selectedSource = null;
      if (sourceId != null) {
        selectedSource = knowledgeSources.find(
          (source) => source.id === sourceId && source.enabled,
        );
        if (!selectedSource) {
          return res.status(400).json({ message: "The selected knowledge source is unavailable." });
        }
      }

      await storage.createMessage({ role: 'user', content: message });

      const sharepointConfig = selectedSource
        ? await storage.getSharepointConfig()
        : null;
      const docs = documentsForSource(
        allDocuments,
        selectedSource,
        sharepointConfig,
        knowledgeSources,
      );
      const appCfg = await storage.getAppSettings();
      const notFoundMessage = appCfg?.notFoundMessage
        ?? "I'm sorry, I couldn't find relevant information for your request in the available documents. Please check your SharePoint library directly or contact your administrator.";
      const rawInstructions = appCfg?.customInstructions?.trim() || null;
      // Strip any not-found fallback the admin may have written in their custom
      // instructions — the Not Found Message field is the sole source of truth.
      const strippedInstructions = stripNotFoundFallback(rawInstructions);
      const customInstructions = preprocessInstructions(strippedInstructions);
      const sourceInstructions = selectedSource?.instructions?.trim()
        ? preprocessInstructions(stripNotFoundFallback(selectedSource.instructions.trim()))
        : null;
      const configuredEscalation = sourceEscalationText(selectedSource);
      const effectiveNotFoundMessage = configuredEscalation || notFoundMessage;

      const context = docs.map(d =>
        `[SOURCE]\nTitle: ${d.title}\nURL: ${d.url}\nType: ${d.type}\nContent: ${d.content}`
      ).join('\n\n---\n\n');

      const systemPrompt = `You are a SharePoint document assistant for this organization.
${customInstructions ? `\nOwner instructions — these are pre-approved by the organisation and govern your tone, style, format, prefix text, and response structure. Follow them precisely for every response:\n${customInstructions}\n` : ''}
${selectedSource ? `\nKnowledge scope — answer only from the selected "${selectedSource.name}" source. Do not use documents from another source.\n` : ''}
${sourceInstructions ? `Source-specific instructions — these apply only to the selected knowledge source:\n${sourceInstructions}\n` : ''}
NOT FOUND OVERRIDE: If the answer is not found in the approved documents below, you MUST respond with EXACTLY the following message — no more, no less. This overrides any alternative not-found or fallback wording that may appear in the owner instructions above:
${effectiveNotFoundMessage}

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

  app.get("/api/settings/public", async (_req, res) => {
    try {
      const s = await storage.getAppSettings();
      res.json({
        assistantName: s?.assistantName ?? "inSite Assistant",
        welcomeMessage: s?.welcomeMessage ?? "Ask me anything about Human Resources or Paid Time Off.",
        responseStyle: extractGlobalResponseStyle(s?.customInstructions ?? ""),
      });
    } catch (err) {
      console.error("GET /api/settings/public error:", err);
      res.status(500).json({ message: err?.message ?? "Internal server error" });
    }
  });

  app.get("/api/settings", requireKnowledgeSourceAdmin, async (req, res) => {
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
        escape(htmlToPlainText(log.assistantResponse)),
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

  app.post("/api/settings", requireKnowledgeSourceAdmin, async (req, res) => {
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
    let syncedSources = 0;
    let failedSources = 0;

    try {
      const knowledgeSources = (await storage.getKnowledgeSources())
        .filter((source) => source.enabled && !source.isPortalWide && source.libraryName);
      if (knowledgeSources.length === 0) {
        return res.status(400).json({
          message: "Enable at least one named knowledge source before syncing.",
        });
      }

      for (const source of knowledgeSources) {
        const sourceConfig = { ...config, libraryName: source.libraryName };
        try {
          const items = await fetchLibraryItems(sourceConfig);
          console.log(`SharePoint sync: found ${items.length} files in "${source.libraryName}"`);

          const importedDocuments = [];
          for (const item of items) {
            try {
              const docWithContent = await fetchDocumentContent(sourceConfig, item);
              importedDocuments.push({
                title: docWithContent.title,
                content: docWithContent.content,
                type: "document",
                url: docWithContent.url,
                source: knowledgeSourceDocumentKey(source),
              });
            } catch (err) {
              console.error(`Failed to sync "${item.title}" from "${source.name}":`, err?.message);
              failed++;
            }
          }

          // Replace only this source after its library was listed successfully,
          // so one failing source cannot erase another source's indexed content.
          await storage.deleteDocumentsBySource(knowledgeSourceDocumentKey(source));
          for (const document of importedDocuments) {
            await storage.createDocument(document);
            synced++;
          }
          syncedSources++;
        } catch (err) {
          failedSources++;
          console.error(`Failed to sync source "${source.name}":`, err?.message);
        }
      }

      if (syncedSources > 0) {
        // Remove the pre-source-model corpus only after at least one named source
        // has been imported successfully.
        await storage.deleteDocumentsBySource("sharepoint");
        await storage.updateSharepointSyncTime();
      }

      res.json({
        synced,
        failed,
        message: `Sync complete. ${synced} document${synced !== 1 ? 's' : ''} imported across ${syncedSources} source${syncedSources !== 1 ? 's' : ''}${failed > 0 ? `, ${failed} document${failed !== 1 ? 's' : ''} failed` : ''}${failedSources > 0 ? `, ${failedSources} source${failedSources !== 1 ? 's' : ''} unavailable` : ''}.`,
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

  async function seedKnowledgeSources() {
    const existing = await storage.getKnowledgeSources();
    if (existing.length === 0) {
      for (const source of DEFAULT_KNOWLEDGE_SOURCES) {
        await storage.createKnowledgeSource(source);
      }

      const sharepointConfig = await storage.getSharepointConfig();
      const configuredLibrary = sharepointConfig?.libraryName?.trim();
      const isExampleLibrary = DEFAULT_KNOWLEDGE_SOURCES.some(
        (source) =>
          source.libraryName
          && normaliseSourceValue(source.libraryName) === normaliseSourceValue(configuredLibrary),
      );
      if (configuredLibrary && !isExampleLibrary) {
        await storage.createKnowledgeSource({
          name: configuredLibrary,
          libraryName: configuredLibrary,
          description: `Searches the ${configuredLibrary} SharePoint library.`,
          instructions: "",
          smeTeam: "",
          contactMethod: "",
          contactDetails: "",
          escalationMessage: "",
          enabled: true,
          isPortalWide: false,
        });
      }
      return;
    }

    // The aggregate scope is a permanent system option. Add it for an
    // installation that already has custom named sources but predates it.
    if (!existing.some((source) => source.isPortalWide)) {
      await storage.createKnowledgeSource(DEFAULT_KNOWLEDGE_SOURCES[0]);
    }
  }

  await seedDatabase();
  await seedKnowledgeSources();

  return httpServer;
}
