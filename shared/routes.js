import { z } from 'zod';
import { insertDocumentSchema, insertSharepointConfigSchema } from './schema.js';

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
      method: 'GET',
      path: '/api/documents',
      responses: {
        200: z.array(z.any()),
      },
    },
    create: {
      method: 'POST',
      path: '/api/documents',
      input: insertDocumentSchema,
      responses: {
        201: z.any(),
        400: errorSchemas.validation,
      },
    },
    delete: {
      method: 'DELETE',
      path: '/api/documents/:id',
      responses: {
        204: z.void(),
        404: errorSchemas.notFound,
      },
    },
  },
  chat: {
    history: {
      method: 'GET',
      path: '/api/chat',
      responses: {
        200: z.array(z.any()),
      },
    },
    send: {
      method: 'POST',
      path: '/api/chat',
      input: z.object({
        message: z.string(),
        username: z.string().optional(),
        sessionId: z.string().optional(),
      }),
      responses: {
        200: z.any(),
        500: errorSchemas.internal,
      },
    },
    clear: {
      method: 'DELETE',
      path: '/api/chat',
      responses: {
        204: z.void(),
      },
    },
  },
  sharepoint: {
    getConfig: {
      method: 'GET',
      path: '/api/sharepoint/config',
      responses: {
        200: z.any(),
      },
    },
    saveConfig: {
      method: 'POST',
      path: '/api/sharepoint/config',
      input: insertSharepointConfigSchema,
      responses: {
        200: z.any(),
        400: errorSchemas.validation,
      },
    },
    testConnection: {
      method: 'POST',
      path: '/api/sharepoint/test',
      responses: {
        200: z.object({ success: z.boolean(), message: z.string() }),
      },
    },
    sync: {
      method: 'POST',
      path: '/api/sharepoint/sync',
      responses: {
        200: z.object({ synced: z.number(), failed: z.number(), message: z.string() }),
        400: z.object({ message: z.string() }),
        500: errorSchemas.internal,
      },
    },
  },
};

export function buildUrl(path, params) {
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
