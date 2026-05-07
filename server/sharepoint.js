import { createRequire } from "module";
const require = createRequire(import.meta.url);

// ─── NTLM helpers (on-premises) ──────────────────────────────────────────────

async function ntlmRequest(options) {
  const httpntlm = await import("httpntlm");
  const client = httpntlm.default || httpntlm;

  return new Promise((resolve, reject) => {
    client.get(options, (err, res) => {
      if (err) return reject(err);
      resolve({ statusCode: res.statusCode, body: res.body });
    });
  });
}

function buildNtlmOptions(url, config, binary = false) {
  return {
    url,
    username: config.username,
    password: config.password,
    domain: config.domain,
    workstation: "",
    binary,
    rejectUnauthorized: !config.allowSelfSigned,
  };
}

// ─── Microsoft Graph helpers (SharePoint Online) ──────────────────────────────

function getOnlineCredentials(config) {
  const tenantId     = process.env.SHAREPOINT_TENANT_ID     || config.tenantId;
  const clientId     = process.env.SHAREPOINT_CLIENT_ID     || config.clientId;
  const clientSecret = process.env.SHAREPOINT_CLIENT_SECRET || config.clientSecret;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      "SharePoint Online credentials are not configured. Set SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, and SHAREPOINT_CLIENT_SECRET as environment variables, or enter them in the Settings page."
    );
  }
  return { tenantId, clientId, clientSecret };
}

async function getGraphToken(config) {
  const { tenantId, clientId, clientSecret } = getOnlineCredentials(config);
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`OAuth token request failed (HTTP ${res.status}): ${text}`);
  }

  const data = await res.json();
  return data.access_token;
}

async function graphRequest(path, token, { binary = false } = {}) {
  const url = path.startsWith("https://") ? path : `https://graph.microsoft.com/v1.0${path}`;

  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  let body;
  if (binary) {
    const buf = await res.arrayBuffer();
    body = Buffer.from(buf);
  } else {
    body = await res.text();
  }

  if (res.status >= 400) {
    console.log(`[Graph API] HTTP ${res.status} from: ${url}`);
    console.log(`[Graph API] Response: ${typeof body === "string" ? body.slice(0, 500) : "(binary)"}`);
  }

  return { statusCode: res.status, body };
}

// Resolve a SharePoint site URL to a Graph site object
async function resolveSite(siteUrl, token) {
  const url = new URL(normaliseSiteUrl(siteUrl));
  const hostname = url.hostname;
  const path = url.pathname && url.pathname !== "/" ? url.pathname : null;

  const endpoint = path
    ? `/sites/${hostname}:${path}`
    : `/sites/${hostname}:/`;

  const res = await graphRequest(endpoint, token);
  if (res.statusCode !== 200) {
    let detail = "";
    try { detail = JSON.parse(res.body)?.error?.message ?? ""; } catch {}
    throw new Error(
      `Could not find SharePoint site (HTTP ${res.statusCode})${detail ? `: ${detail}` : ""}. Check the Site URL.`
    );
  }
  return JSON.parse(res.body);
}

// Find the drive (document library) by name within a site
async function resolveDrive(siteId, libraryName, token) {
  const res = await graphRequest(`/sites/${siteId}/drives`, token);
  if (res.statusCode !== 200) {
    throw new Error(`Could not list document libraries (HTTP ${res.statusCode}).`);
  }
  const drives = JSON.parse(res.body).value ?? [];
  const drive = drives.find(
    (d) => d.name.toLowerCase() === libraryName.toLowerCase()
  );
  if (!drive) {
    const available = drives.map((d) => d.name).join(", ");
    throw new Error(
      `Library "${libraryName}" not found. Available libraries: ${available || "(none)"}`
    );
  }
  return drive;
}

// ─── Shared helpers ───────────────────────────────────────────────────────────

export function normaliseSiteUrl(raw) {
  return raw.trim().replace(/\/+$/, "");
}

async function extractTextContent(buffer, extension, fileName) {
  switch (extension) {
    case "txt":
    case "csv":
    case "md":
      return buffer.toString("utf-8").slice(0, 8000);

    case "docx":
    case "doc": {
      const mammoth = await import("mammoth");
      const result = await (mammoth.default || mammoth).extractRawText({ buffer });
      return (result.value ?? "").slice(0, 8000);
    }

    case "pdf": {
      const { PDFParse } = require("pdf-parse");
      const parser = new PDFParse({ data: buffer });
      const result = await parser.getText();
      return (result.text ?? "").slice(0, 8000);
    }

    default:
      return `[File: ${fileName} — content extraction not supported for .${extension} files]`;
  }
}

// ─── Exported functions ───────────────────────────────────────────────────────

export async function fetchLibraryItems(config) {
  // ── SharePoint Online via Microsoft Graph ──────────────────────────────────
  if (config.mode === "online") {
    const token = await getGraphToken(config);
    const site  = await resolveSite(config.siteUrl, token);
    const drive = await resolveDrive(site.id, config.libraryName, token);

    const res = await graphRequest(`/drives/${drive.id}/root/children?$top=500`, token);
    if (res.statusCode !== 200) {
      throw new Error(`Could not list files in library (HTTP ${res.statusCode}).`);
    }

    const items = JSON.parse(res.body).value ?? [];
    return items
      .filter((item) => item.file) // exclude folders
      .map((item) => {
        const ext = item.name.split(".").pop()?.toLowerCase() ?? "";
        const title = item.name.replace(/\.[^.]+$/, "");
        return {
          title,
          url: item.webUrl,
          fileRef: item.name,
          extension: ext,
          _driveId: drive.id,
          _itemId: item.id,
        };
      });
  }

  // ── On-premises via NTLM ───────────────────────────────────────────────────
  const siteUrl = normaliseSiteUrl(config.siteUrl);
  const apiUrl =
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(config.libraryName)}')/items` +
    `?$select=Title,FileLeafRef,FileRef,File_x0020_Type&$filter=FSObjType eq 0&$top=500`;

  const res = await ntlmRequest({
    ...buildNtlmOptions(apiUrl, config),
    headers: { Accept: "application/json;odata=verbose" },
  });

  if (res.statusCode !== 200) {
    throw new Error(
      `SharePoint returned status ${res.statusCode}. Check your credentials and site URL.`
    );
  }

  let data;
  try {
    data = JSON.parse(res.body);
  } catch {
    throw new Error("Could not parse SharePoint response. Verify the site URL and library name.");
  }

  const items = data?.d?.results ?? data?.value ?? [];
  return items
    .filter((item) => item.FileRef && item.FileLeafRef)
    .map((item) => {
      const fileRef = item.FileRef;
      const fileName = item.FileLeafRef;
      const ext = fileName.split(".").pop()?.toLowerCase() ?? "";
      const title = item.Title || fileName.replace(/\.[^.]+$/, "");
      const fullUrl = `${new URL(siteUrl).origin}${fileRef}`;
      return { title, url: fullUrl, fileRef, extension: ext };
    });
}

export async function fetchDocumentContent(config, doc) {
  const supportedExtensions = ["txt", "csv", "md", "docx", "doc", "pdf"];

  if (!supportedExtensions.includes(doc.extension)) {
    return {
      ...doc,
      content: `[File: ${doc.title} — content extraction not supported for .${doc.extension} files]`,
    };
  }

  // ── SharePoint Online via Microsoft Graph ──────────────────────────────────
  if (config.mode === "online") {
    if (!doc._driveId || !doc._itemId) {
      throw new Error(`Missing Graph identifiers for "${doc.title}". Please re-sync documents.`);
    }
    const token = await getGraphToken(config);
    const res = await graphRequest(
      `/drives/${doc._driveId}/items/${doc._itemId}/content`,
      token,
      { binary: true }
    );
    if (res.statusCode !== 200) {
      throw new Error(`Failed to fetch content for "${doc.title}" (HTTP ${res.statusCode}).`);
    }
    const buffer = Buffer.isBuffer(res.body) ? res.body : Buffer.from(res.body);
    const content = await extractTextContent(buffer, doc.extension, doc.title);
    return { ...doc, content };
  }

  // ── On-premises via NTLM ───────────────────────────────────────────────────
  const siteUrl = normaliseSiteUrl(config.siteUrl);
  const fileUrl =
    `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(doc.fileRef)}')/$value`;

  const res = await ntlmRequest(buildNtlmOptions(fileUrl, config, true));

  if (res.statusCode !== 200) {
    throw new Error(`Failed to fetch file content for ${doc.title} (status ${res.statusCode})`);
  }

  const buffer = Buffer.isBuffer(res.body) ? res.body : Buffer.from(res.body);
  const content = await extractTextContent(buffer, doc.extension, doc.title);
  return { ...doc, content };
}

export async function testSharepointConnection(config) {
  try {
    // ── SharePoint Online via Microsoft Graph ────────────────────────────────
    if (config.mode === "online") {
      const token = await getGraphToken(config);
      const site  = await resolveSite(config.siteUrl, token);
      const name  = site.displayName ?? site.name ?? "SharePoint Online";
      return { success: true, message: `Connected to "${name}"` };
    }

    // ── On-premises via NTLM ─────────────────────────────────────────────────
    const siteUrl = normaliseSiteUrl(config.siteUrl);
    const apiUrl  = `${siteUrl}/_api/web/title`;

    const res = await ntlmRequest({
      ...buildNtlmOptions(apiUrl, { ...config, siteUrl }),
      headers: { Accept: "application/json;odata=verbose" },
    });

    if (res.statusCode === 200) {
      let title = "SharePoint";
      try {
        const data = JSON.parse(res.body);
        title = data?.d?.Title ?? data?.value ?? "SharePoint";
      } catch {}
      return { success: true, message: `Connected to "${title}"` };
    } else if (res.statusCode === 401 || res.statusCode === 403) {
      return {
        success: false,
        message: `Authentication failed (HTTP ${res.statusCode}). Check your domain, username, and password.`,
      };
    } else if (res.statusCode === 404) {
      return {
        success: false,
        message: `Site not found (HTTP 404). Tried: ${apiUrl}\n\nCommon fixes:\n• Remove any trailing slash from the URL\n• Use the root site URL only\n• Do not include page names or document library paths`,
      };
    } else {
      return { success: false, message: `Server returned HTTP ${res.statusCode}. Tried: ${apiUrl}` };
    }
  } catch (err) {
    const msg = err?.message ?? "Unknown error";
    if (msg.includes("ECONNREFUSED")) {
      return { success: false, message: `Connection refused. Make sure SharePoint is reachable from this server.` };
    }
    if (msg.includes("ENOTFOUND")) {
      return { success: false, message: `Host not found. Check the hostname is correct and DNS is resolving.` };
    }
    if (msg.includes("certificate") || msg.includes("SSL") || msg.includes("self signed")) {
      return { success: false, message: `SSL certificate error. Enable "Allow self-signed certificates" in settings if using an internal certificate.` };
    }
    if (msg.includes("OAuth token")) {
      return { success: false, message: `OAuth authentication failed: ${msg}` };
    }
    return { success: false, message: `Connection error: ${msg}` };
  }
}
