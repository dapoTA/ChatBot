import { getPool, sql } from "./db.mssql.js";
import { encrypt, decrypt } from "./crypto.js";

const ENCRYPTED_FIELDS = ["password", "clientSecret"];

function mapDocument(row) {
  if (!row) return null;
  return {
    id: row.id,
    title: row.title,
    content: row.content,
    type: row.type,
    url: row.url,
    source: row.source,
    createdAt: row.created_at,
  };
}

function mapMessage(row) {
  if (!row) return null;
  return {
    id: row.id,
    role: row.role,
    content: row.content,
    createdAt: row.created_at,
  };
}

function mapSharepointConfig(row) {
  if (!row) return null;
  return {
    id: row.id,
    mode: row.mode,
    siteUrl: row.site_url,
    siteUrlOnprem: row.site_url_onprem,
    siteUrlOnline: row.site_url_online,
    libraryName: row.library_name,
    domain: row.domain,
    username: row.username,
    password: row.password,
    allowSelfSigned: row.allow_self_signed === true || row.allow_self_signed === 1,
    tenantId: row.tenant_id,
    clientId: row.client_id,
    clientSecret: row.client_secret,
    lastSyncedAt: row.last_synced_at,
    updatedAt: row.updated_at,
  };
}

function mapAppSettings(row) {
  if (!row) return null;
  return {
    id: row.id,
    assistantName: row.assistant_name,
    welcomeMessage: row.welcome_message,
    notFoundMessage: row.not_found_message,
    customInstructions: row.custom_instructions,
    temperature: row.temperature,
    topP: row.top_p,
    maxTokens: row.max_tokens,
    frequencyPenalty: row.frequency_penalty,
    presencePenalty: row.presence_penalty,
    updatedAt: row.updated_at,
  };
}

export class DatabaseStorage {
  async getDocuments() {
    const pool = await getPool();
    const result = await pool.request().query(
      "SELECT id, title, content, type, url, source, created_at FROM documents ORDER BY created_at DESC"
    );
    return result.recordset.map(mapDocument);
  }

  async createDocument(doc) {
    const pool = await getPool();
    const result = await pool.request()
      .input("title", sql.NVarChar(sql.MAX), doc.title)
      .input("content", sql.NVarChar(sql.MAX), doc.content)
      .input("type", sql.NVarChar(255), doc.type)
      .input("url", sql.NVarChar(sql.MAX), doc.url)
      .input("source", sql.NVarChar(255), doc.source ?? "manual")
      .query(`INSERT INTO documents (title, content, type, url, source)
              OUTPUT INSERTED.*
              VALUES (@title, @content, @type, @url, @source)`);
    return mapDocument(result.recordset[0]);
  }

  async deleteDocument(id) {
    const pool = await getPool();
    await pool.request()
      .input("id", sql.Int, id)
      .query("DELETE FROM documents WHERE id = @id");
  }

  async deleteDocumentsBySource(source) {
    const pool = await getPool();
    await pool.request()
      .input("source", sql.NVarChar(255), source)
      .query("DELETE FROM documents WHERE source = @source");
  }

  async getMessages() {
    const pool = await getPool();
    const result = await pool.request().query(
      "SELECT id, role, content, created_at FROM messages ORDER BY created_at ASC"
    );
    return result.recordset.map(mapMessage);
  }

  async createMessage(msg) {
    const pool = await getPool();
    const result = await pool.request()
      .input("role", sql.NVarChar(50), msg.role)
      .input("content", sql.NVarChar(sql.MAX), msg.content)
      .query(`INSERT INTO messages (role, content)
              OUTPUT INSERTED.*
              VALUES (@role, @content)`);
    return mapMessage(result.recordset[0]);
  }

  async clearMessages() {
    const pool = await getPool();
    await pool.request().query("DELETE FROM messages");
  }

  async getSharepointConfig() {
    const pool = await getPool();
    const result = await pool.request().query(
      "SELECT TOP 1 * FROM sharepoint_configs"
    );
    const row = result.recordset[0] ?? null;
    if (!row) return null;
    const mapped = mapSharepointConfig(row);
    for (const field of ENCRYPTED_FIELDS) {
      if (mapped[field]) mapped[field] = decrypt(mapped[field]);
    }
    return mapped;
  }

  async upsertSharepointConfig(config) {
    const pool = await getPool();

    const toStore = { ...config };
    for (const field of ENCRYPTED_FIELDS) {
      if (toStore[field] && toStore[field] !== "••••••••") {
        toStore[field] = encrypt(toStore[field]);
      }
    }

    const existing = await pool.request().query("SELECT TOP 1 id FROM sharepoint_configs");
    const existingRow = existing.recordset[0] ?? null;

    if (existingRow) {
      await pool.request()
        .input("id", sql.Int, existingRow.id)
        .input("mode", sql.NVarChar(50), toStore.mode ?? "onprem")
        .input("site_url", sql.NVarChar(sql.MAX), toStore.siteUrl ?? "")
        .input("site_url_onprem", sql.NVarChar(sql.MAX), toStore.siteUrlOnprem ?? null)
        .input("site_url_online", sql.NVarChar(sql.MAX), toStore.siteUrlOnline ?? null)
        .input("library_name", sql.NVarChar(255), toStore.libraryName ?? "Documents")
        .input("domain", sql.NVarChar(255), toStore.domain ?? "")
        .input("username", sql.NVarChar(255), toStore.username ?? "")
        .input("password", sql.NVarChar(sql.MAX), toStore.password ?? "")
        .input("allow_self_signed", sql.Bit, toStore.allowSelfSigned ? 1 : 0)
        .input("tenant_id", sql.NVarChar(255), toStore.tenantId ?? null)
        .input("client_id", sql.NVarChar(255), toStore.clientId ?? null)
        .input("client_secret", sql.NVarChar(sql.MAX), toStore.clientSecret ?? null)
        .query(`UPDATE sharepoint_configs SET
                  mode = @mode, site_url = @site_url,
                  site_url_onprem = @site_url_onprem, site_url_online = @site_url_online,
                  library_name = @library_name, domain = @domain,
                  username = @username, password = @password,
                  allow_self_signed = @allow_self_signed,
                  tenant_id = @tenant_id, client_id = @client_id, client_secret = @client_secret,
                  updated_at = GETDATE()
                WHERE id = @id`);
    } else {
      await pool.request()
        .input("mode", sql.NVarChar(50), toStore.mode ?? "onprem")
        .input("site_url", sql.NVarChar(sql.MAX), toStore.siteUrl ?? "")
        .input("site_url_onprem", sql.NVarChar(sql.MAX), toStore.siteUrlOnprem ?? null)
        .input("site_url_online", sql.NVarChar(sql.MAX), toStore.siteUrlOnline ?? null)
        .input("library_name", sql.NVarChar(255), toStore.libraryName ?? "Documents")
        .input("domain", sql.NVarChar(255), toStore.domain ?? "")
        .input("username", sql.NVarChar(255), toStore.username ?? "")
        .input("password", sql.NVarChar(sql.MAX), toStore.password ?? "")
        .input("allow_self_signed", sql.Bit, toStore.allowSelfSigned ? 1 : 0)
        .input("tenant_id", sql.NVarChar(255), toStore.tenantId ?? null)
        .input("client_id", sql.NVarChar(255), toStore.clientId ?? null)
        .input("client_secret", sql.NVarChar(sql.MAX), toStore.clientSecret ?? null)
        .query(`INSERT INTO sharepoint_configs
                  (mode, site_url, site_url_onprem, site_url_online, library_name,
                   domain, username, password, allow_self_signed,
                   tenant_id, client_id, client_secret)
                VALUES
                  (@mode, @site_url, @site_url_onprem, @site_url_online, @library_name,
                   @domain, @username, @password, @allow_self_signed,
                   @tenant_id, @client_id, @client_secret)`);
    }

    return this.getSharepointConfig();
  }

  async updateSharepointSyncTime() {
    const pool = await getPool();
    const existing = await pool.request().query("SELECT TOP 1 id FROM sharepoint_configs");
    const row = existing.recordset[0] ?? null;
    if (row) {
      await pool.request()
        .input("id", sql.Int, row.id)
        .query("UPDATE sharepoint_configs SET last_synced_at = GETDATE() WHERE id = @id");
    }
  }

  async getAppSettings() {
    const pool = await getPool();
    const result = await pool.request().query("SELECT TOP 1 * FROM app_settings");
    return mapAppSettings(result.recordset[0] ?? null);
  }

  async upsertAppSettings(settings) {
    const pool = await getPool();
    const existing = await pool.request().query("SELECT TOP 1 id FROM app_settings");
    const existingRow = existing.recordset[0] ?? null;

    if (existingRow) {
      await pool.request()
        .input("id", sql.Int, existingRow.id)
        .input("assistant_name", sql.NVarChar(255), settings.assistantName)
        .input("welcome_message", sql.NVarChar(sql.MAX), settings.welcomeMessage)
        .input("not_found_message", sql.NVarChar(sql.MAX), settings.notFoundMessage)
        .input("custom_instructions", sql.NVarChar(sql.MAX), settings.customInstructions ?? null)
        .input("temperature", sql.Float, settings.temperature)
        .input("top_p", sql.Float, settings.topP)
        .input("max_tokens", sql.Int, settings.maxTokens)
        .input("frequency_penalty", sql.Float, settings.frequencyPenalty)
        .input("presence_penalty", sql.Float, settings.presencePenalty)
        .query(`UPDATE app_settings SET
                  assistant_name = @assistant_name,
                  welcome_message = @welcome_message,
                  not_found_message = @not_found_message,
                  custom_instructions = @custom_instructions,
                  temperature = @temperature, top_p = @top_p,
                  max_tokens = @max_tokens,
                  frequency_penalty = @frequency_penalty,
                  presence_penalty = @presence_penalty,
                  updated_at = GETDATE()
                WHERE id = @id`);
    } else {
      await pool.request()
        .input("assistant_name", sql.NVarChar(255), settings.assistantName)
        .input("welcome_message", sql.NVarChar(sql.MAX), settings.welcomeMessage)
        .input("not_found_message", sql.NVarChar(sql.MAX), settings.notFoundMessage)
        .input("custom_instructions", sql.NVarChar(sql.MAX), settings.customInstructions ?? null)
        .input("temperature", sql.Float, settings.temperature)
        .input("top_p", sql.Float, settings.topP)
        .input("max_tokens", sql.Int, settings.maxTokens)
        .input("frequency_penalty", sql.Float, settings.frequencyPenalty)
        .input("presence_penalty", sql.Float, settings.presencePenalty)
        .query(`INSERT INTO app_settings
                  (assistant_name, welcome_message, not_found_message, custom_instructions,
                   temperature, top_p, max_tokens, frequency_penalty, presence_penalty)
                VALUES
                  (@assistant_name, @welcome_message, @not_found_message, @custom_instructions,
                   @temperature, @top_p, @max_tokens, @frequency_penalty, @presence_penalty)`);
    }

    return this.getAppSettings();
  }
}

export const storage = new DatabaseStorage();
