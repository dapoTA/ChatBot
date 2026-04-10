import { createRequire } from "module";
const require = createRequire(import.meta.url);

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

export function normaliseSiteUrl(raw) {
  return raw.trim().replace(/\/+$/, "");
}

export async function fetchLibraryItems(config) {
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

export async function fetchDocumentContent(config, doc) {
  const siteUrl = normaliseSiteUrl(config.siteUrl);
  const supportedExtensions = ["txt", "csv", "md", "docx", "doc", "pdf"];

  if (!supportedExtensions.includes(doc.extension)) {
    return {
      ...doc,
      content: `[File: ${doc.title} — content extraction not supported for .${doc.extension} files]`,
    };
  }

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
  const siteUrl = normaliseSiteUrl(config.siteUrl);
  const apiUrl = `${siteUrl}/_api/web/title`;

  try {
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
        message: `Site not found (HTTP 404). Tried: ${apiUrl}\n\nCommon fixes:\n• Remove any trailing slash from the URL\n• Use the root site URL only (e.g. http://sharepoint.company.com/sites/MySite)\n• Do not include page names or document library paths`,
      };
    } else {
      return {
        success: false,
        message: `Server returned HTTP ${res.statusCode}. Tried: ${apiUrl}`,
      };
    }
  } catch (err) {
    const msg = err?.message ?? "Unknown error";
    if (msg.includes("ECONNREFUSED")) {
      return {
        success: false,
        message: `Connection refused at ${siteUrl}. Make sure SharePoint is reachable from this server and the port is correct.`,
      };
    }
    if (msg.includes("ENOTFOUND")) {
      return {
        success: false,
        message: `Host not found: "${siteUrl}". Check the hostname is correct and DNS is resolving on this server.`,
      };
    }
    if (msg.includes("certificate") || msg.includes("SSL") || msg.includes("self signed")) {
      return {
        success: false,
        message: `SSL certificate error. If your SharePoint uses a self-signed certificate, enable "Allow self-signed certificates" in the connection settings.`,
      };
    }
    return { success: false, message: `Connection error: ${msg}\n\nTried: ${apiUrl}` };
  }
}
