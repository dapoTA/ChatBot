// @ts-nocheck - Third-party packages (httpntlm, mammoth, pdf-parse) lack complete type declarations
import type { SharepointConfig } from "@shared/schema";

interface NtlmOptions {
  url: string;
  username: string;
  password: string;
  domain: string;
  workstation?: string;
  binary?: boolean;
  rejectUnauthorized?: boolean;
}

async function ntlmRequest(options: NtlmOptions): Promise<{ statusCode: number; body: any }> {
  const httpntlm = require("httpntlm");

  return new Promise((resolve, reject) => {
    httpntlm.get(options, (err: Error | null, res: any) => {
      if (err) return reject(err);
      resolve({ statusCode: res.statusCode, body: res.body });
    });
  });
}

function buildNtlmOptions(url: string, config: SharepointConfig, binary = false): NtlmOptions {
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

export interface SPDocument {
  title: string;
  url: string;
  fileRef: string;
  extension: string;
}

export interface SPDocumentWithContent extends SPDocument {
  content: string;
}

// Fetch list of files from a SharePoint document library
export async function fetchLibraryItems(config: SharepointConfig): Promise<SPDocument[]> {
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

  let data: any;
  try {
    data = JSON.parse(res.body as string);
  } catch {
    throw new Error("Could not parse SharePoint response. Verify the site URL and library name.");
  }

  const items: any[] = data?.d?.results ?? data?.value ?? [];

  return items
    .filter((item: any) => item.FileRef && item.FileLeafRef)
    .map((item: any) => {
      const fileRef: string = item.FileRef;
      const fileName: string = item.FileLeafRef;
      const ext = fileName.split(".").pop()?.toLowerCase() ?? "";
      const title = item.Title || fileName.replace(/\.[^.]+$/, "");
      const fullUrl = `${siteUrl}${fileRef}`;
      return { title, url: fullUrl, fileRef, extension: ext };
    });
}

async function extractTextContent(buffer: Buffer, extension: string, fileName: string): Promise<string> {
  switch (extension) {
    case "txt":
    case "csv":
    case "md":
      return buffer.toString("utf-8").slice(0, 8000);

    case "docx":
    case "doc": {
      const mammoth = require("mammoth");
      const result = await mammoth.extractRawText({ buffer });
      return (result.value ?? "").slice(0, 8000);
    }

    case "pdf": {
      const pdfParse = require("pdf-parse");
      const data = await pdfParse(buffer);
      return (data.text ?? "").slice(0, 8000);
    }

    default:
      return `[File: ${fileName} — content extraction not supported for .${extension} files]`;
  }
}

// Fetch and extract content for a single file
export async function fetchDocumentContent(
  config: SharepointConfig,
  doc: SPDocument
): Promise<SPDocumentWithContent> {
  const supportedExtensions = ["txt", "csv", "md", "docx", "doc", "pdf"];

  if (!supportedExtensions.includes(doc.extension)) {
    return {
      ...doc,
      content: `[File: ${doc.title} — content extraction not supported for .${doc.extension} files]`,
    };
  }

  const fileUrl =
    `${config.siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(doc.fileRef)}')/$value`;

  const res = await ntlmRequest(buildNtlmOptions(fileUrl, config, true));

  if (res.statusCode !== 200) {
    throw new Error(`Failed to fetch file content for ${doc.title} (status ${res.statusCode})`);
  }

  const buffer = Buffer.isBuffer(res.body) ? res.body : Buffer.from(res.body as string);
  const content = await extractTextContent(buffer, doc.extension, doc.title);

  return { ...doc, content };
}

// Normalise site URL — strip trailing slash, ensure no double-slashes
export function normaliseSiteUrl(raw: string): string {
  return raw.trim().replace(/\/+$/, "");
}

// Test the SharePoint connection
export async function testSharepointConnection(
  config: SharepointConfig
): Promise<{ success: boolean; message: string }> {
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
        const data = JSON.parse(res.body as string);
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
  } catch (err: any) {
    const msg: string = err?.message ?? "Unknown error";
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
