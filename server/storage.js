import { pool, sql } from "./db.js";

function row2doc(r) {
  return { id: r.id, title: r.title, content: r.content, type: r.type, url: r.url, source: r.source, createdAt: r.created_at };
}
function row2msg(r) {
  return { id: r.id, role: r.role, content: r.content, createdAt: r.created_at };
}
function row2spConfig(r) {
  return {
    id: r.id, siteUrl: r.site_url, domain: r.domain, username: r.username,
    password: r.password, libraryName: r.library_name,
    allowSelfSigned: Boolean(r.allow_self_signed),
    lastSyncedAt: r.last_synced_at, updatedAt: r.updated_at,
  };
}
function row2settings(r) {
  return {
    id: r.id, assistantName: r.assistant_name, welcomeMessage: r.welcome_message,
    notFoundMessage: r.not_found_message, customInstructions: r.custom_instructions,
    temperature: r.temperature, topP: r.top_p, maxTokens: r.max_tokens,
    frequencyPenalty: r.frequency_penalty, presencePenalty: r.presence_penalty,
    updatedAt: r.updated_at,
  };
}

export class DatabaseStorage {
  async getDocuments() {
    const res = await pool.request().query("SELECT * FROM documents ORDER BY created_at DESC");
    return res.recordset.map(row2doc);
  }

  async createDocument(doc) {
    const res = await pool.request()
      .input("title",   sql.NVarChar(sql.MAX), doc.title)
      .input("content", sql.NVarChar(sql.MAX), doc.content)
      .input("type",    sql.NVarChar(sql.MAX), doc.type)
      .input("url",     sql.NVarChar(sql.MAX), doc.url)
      .input("source",  sql.NVarChar(sql.MAX), doc.source ?? "manual")
      .query(`INSERT INTO documents (title,content,type,url,source) OUTPUT INSERTED.* VALUES (@title,@content,@type,@url,@source)`);
    return row2doc(res.recordset[0]);
  }

  async deleteDocument(id) {
    await pool.request().input("id", sql.Int, id).query("DELETE FROM documents WHERE id=@id");
  }

  async deleteDocumentsBySource(source) {
    await pool.request().input("source", sql.NVarChar(sql.MAX), source).query("DELETE FROM documents WHERE source=@source");
  }

  async getMessages() {
    const res = await pool.request().query("SELECT * FROM messages ORDER BY created_at ASC");
    return res.recordset.map(row2msg);
  }

  async createMessage(msg) {
    const res = await pool.request()
      .input("role",    sql.NVarChar(sql.MAX), msg.role)
      .input("content", sql.NVarChar(sql.MAX), msg.content)
      .query(`INSERT INTO messages (role,content) OUTPUT INSERTED.* VALUES (@role,@content)`);
    return row2msg(res.recordset[0]);
  }

  async clearMessages() {
    await pool.request().query("DELETE FROM messages");
  }

  async getSharepointConfig() {
    const res = await pool.request().query("SELECT TOP 1 * FROM sharepoint_configs");
    return res.recordset.length ? row2spConfig(res.recordset[0]) : null;
  }

  async upsertSharepointConfig(config) {
    const existing = await this.getSharepointConfig();
    if (existing) {
      const res = await pool.request()
        .input("id",              sql.Int,             existing.id)
        .input("siteUrl",         sql.NVarChar(sql.MAX), config.siteUrl)
        .input("domain",          sql.NVarChar(sql.MAX), config.domain)
        .input("username",        sql.NVarChar(sql.MAX), config.username)
        .input("password",        sql.NVarChar(sql.MAX), config.password)
        .input("libraryName",     sql.NVarChar(sql.MAX), config.libraryName)
        .input("allowSelfSigned", sql.Bit,             config.allowSelfSigned ? 1 : 0)
        .query(`
          UPDATE sharepoint_configs
          SET site_url=@siteUrl, domain=@domain, username=@username, password=@password,
              library_name=@libraryName, allow_self_signed=@allowSelfSigned, updated_at=GETDATE()
          OUTPUT INSERTED.*
          WHERE id=@id`);
      return row2spConfig(res.recordset[0]);
    } else {
      const res = await pool.request()
        .input("siteUrl",         sql.NVarChar(sql.MAX), config.siteUrl)
        .input("domain",          sql.NVarChar(sql.MAX), config.domain)
        .input("username",        sql.NVarChar(sql.MAX), config.username)
        .input("password",        sql.NVarChar(sql.MAX), config.password)
        .input("libraryName",     sql.NVarChar(sql.MAX), config.libraryName)
        .input("allowSelfSigned", sql.Bit,             config.allowSelfSigned ? 1 : 0)
        .query(`
          INSERT INTO sharepoint_configs (site_url,domain,username,password,library_name,allow_self_signed)
          OUTPUT INSERTED.*
          VALUES (@siteUrl,@domain,@username,@password,@libraryName,@allowSelfSigned)`);
      return row2spConfig(res.recordset[0]);
    }
  }

  async updateSharepointSyncTime() {
    const existing = await this.getSharepointConfig();
    if (existing) {
      await pool.request().input("id", sql.Int, existing.id)
        .query("UPDATE sharepoint_configs SET last_synced_at=GETDATE() WHERE id=@id");
    }
  }

  async getAppSettings() {
    const res = await pool.request().query("SELECT TOP 1 * FROM app_settings");
    return res.recordset.length ? row2settings(res.recordset[0]) : null;
  }

  async upsertAppSettings(settings) {
    const existing = await this.getAppSettings();
    if (existing) {
      const res = await pool.request()
        .input("id",               sql.Int,             existing.id)
        .input("assistantName",    sql.NVarChar(sql.MAX), settings.assistantName)
        .input("welcomeMessage",   sql.NVarChar(sql.MAX), settings.welcomeMessage)
        .input("notFoundMessage",  sql.NVarChar(sql.MAX), settings.notFoundMessage)
        .input("customInstructions", sql.NVarChar(sql.MAX), settings.customInstructions ?? null)
        .input("temperature",      sql.Real,            settings.temperature)
        .input("topP",             sql.Real,            settings.topP)
        .input("maxTokens",        sql.Int,             settings.maxTokens)
        .input("frequencyPenalty", sql.Real,            settings.frequencyPenalty)
        .input("presencePenalty",  sql.Real,            settings.presencePenalty)
        .query(`
          UPDATE app_settings
          SET assistant_name=@assistantName, welcome_message=@welcomeMessage,
              not_found_message=@notFoundMessage, custom_instructions=@customInstructions,
              temperature=@temperature, top_p=@topP, max_tokens=@maxTokens,
              frequency_penalty=@frequencyPenalty, presence_penalty=@presencePenalty, updated_at=GETDATE()
          OUTPUT INSERTED.*
          WHERE id=@id`);
      return row2settings(res.recordset[0]);
    } else {
      const res = await pool.request()
        .input("assistantName",    sql.NVarChar(sql.MAX), settings.assistantName)
        .input("welcomeMessage",   sql.NVarChar(sql.MAX), settings.welcomeMessage)
        .input("notFoundMessage",  sql.NVarChar(sql.MAX), settings.notFoundMessage)
        .input("customInstructions", sql.NVarChar(sql.MAX), settings.customInstructions ?? null)
        .input("temperature",      sql.Real,            settings.temperature)
        .input("topP",             sql.Real,            settings.topP)
        .input("maxTokens",        sql.Int,             settings.maxTokens)
        .input("frequencyPenalty", sql.Real,            settings.frequencyPenalty)
        .input("presencePenalty",  sql.Real,            settings.presencePenalty)
        .query(`
          INSERT INTO app_settings
            (assistant_name,welcome_message,not_found_message,custom_instructions,
             temperature,top_p,max_tokens,frequency_penalty,presence_penalty)
          OUTPUT INSERTED.*
          VALUES (@assistantName,@welcomeMessage,@notFoundMessage,@customInstructions,
                  @temperature,@topP,@maxTokens,@frequencyPenalty,@presencePenalty)`);
      return row2settings(res.recordset[0]);
    }
  }
}

export const storage = new DatabaseStorage();
