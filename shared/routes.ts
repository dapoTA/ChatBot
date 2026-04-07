import { z } from 'zod';
import { insertDocumentSchema, insertSharepointConfigSchema, documents, messages, sharepointConfigs } from './schema';

export const errorSchemas = {
  validation: z.object({
    message: z.string(),
    field: z.string().optional(),
  }),
  notFound: z.object({
    message: z.string(),
  }),
  internal: z.object({
    message: z.string(),
  }),
};

export const api = {
  documents: {
    list: {
      method: 'GET' as const,
      path: '/api/documents',
      responses: {
        200: z.array(z.custom<typeof documents.$inferSelect>()),
      },
    },
    create: {
      method: 'POST' as const,
      path: '/api/documents',
      input: insertDocumentSchema,
      responses: {
        201: z.custom<typeof documents.$inferSelect>(),
        400: errorSchemas.validation,
      },
    },
    delete: {
      method: 'DELETE' as const,
      path: '/api/documents/:id',
      responses: {
        204: z.void(),
        404: errorSchemas.notFound,
      },
    },
  },
  chat: {
    history: {
      method: 'GET' as const,
      path: '/api/chat',
      responses: {
        200: z.array(z.custom<typeof messages.$inferSelect>()),
      },
    },
    send: {
      method: 'POST' as const,
      path: '/api/chat',
      input: z.object({ message: z.string() }),
      responses: {
        200: z.custom<typeof messages.$inferSelect>(),
        500: errorSchemas.internal,
      },
    },
    clear: {
      method: 'DELETE' as const,
      path: '/api/chat',
      responses: {
        204: z.void(),
      },
    },
  },
  sharepoint: {
    getConfig: {
      method: 'GET' as const,
      path: '/api/sharepoint/config',
      responses: {
        200: z.custom<typeof sharepointConfigs.$inferSelect | null>(),
      },
    },
    saveConfig: {
      method: 'POST' as const,
      path: '/api/sharepoint/config',
      input: insertSharepointConfigSchema,
      responses: {
        200: z.custom<typeof sharepointConfigs.$inferSelect>(),
        400: errorSchemas.validation,
      },
    },
    testConnection: {
      method: 'POST' as const,
      path: '/api/sharepoint/test',
      responses: {
        200: z.object({ success: z.boolean(), message: z.string() }),
      },
    },
    sync: {
      method: 'POST' as const,
      path: '/api/sharepoint/sync',
      responses: {
        200: z.object({ synced: z.number(), failed: z.number(), message: z.string() }),
        400: z.object({ message: z.string() }),
        500: errorSchemas.internal,
      },
    },
  },
};

export function buildUrl(path: string, params?: Record<string, string | number>): string {
  let url = path;
  if (params) {
    Object.entries(params).forEach(([key, value]) => {
      if (url.includes(`:${key}`)) {
        url = url.replace(`:${key}`, String(value));
      }
    });
  }
  return url;
}

export type DocumentInput = z.infer<typeof api.documents.create.input>;
export type ChatInput = z.infer<typeof api.chat.send.input>;
export type SharepointConfigInput = z.infer<typeof api.sharepoint.saveConfig.input>;
