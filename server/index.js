#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { OutlookAuthManager } from './auth/auth.js';
import { 
  authenticateTool,
  listEmailsTool,
  sendEmailTool,
  listEventsTool,
  createEventTool
} from './tools/index.js';

const server = new Server(
  {
    name: 'outlook-mcp',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

const authManager = new OutlookAuthManager(
  process.env.AZURE_CLIENT_ID,
  process.env.AZURE_TENANT_ID
);

server.setRequestHandler('tools/list', async () => {
  return {
    tools: [
      {
        name: 'outlook_authenticate',
        description: 'Authenticate with Microsoft Outlook using OAuth 2.0',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'outlook_list_emails',
        description: 'List emails from Outlook inbox or specified folder',
        inputSchema: {
          type: 'object',
          properties: {
            folder: {
              type: 'string',
              description: 'Folder to list emails from (default: inbox)',
              default: 'inbox',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of emails to return',
              default: 10,
            },
            filter: {
              type: 'string',
              description: 'OData filter query for emails',
            },
          },
        },
      },
      {
        name: 'outlook_send_email',
        description: 'Send an email through Outlook',
        inputSchema: {
          type: 'object',
          properties: {
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Recipient email addresses',
            },
            subject: {
              type: 'string',
              description: 'Email subject',
            },
            body: {
              type: 'string',
              description: 'Email body content',
            },
            bodyType: {
              type: 'string',
              enum: ['text', 'html'],
              default: 'text',
              description: 'Body content type',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC recipients',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC recipients',
            },
          },
          required: ['to', 'subject', 'body'],
        },
      },
      {
        name: 'outlook_list_events',
        description: 'List calendar events from Outlook',
        inputSchema: {
          type: 'object',
          properties: {
            startDateTime: {
              type: 'string',
              description: 'Start date/time in ISO 8601 format',
            },
            endDateTime: {
              type: 'string',
              description: 'End date/time in ISO 8601 format',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of events to return',
              default: 10,
            },
            calendar: {
              type: 'string',
              description: 'Calendar ID (default: primary calendar)',
            },
          },
        },
      },
      {
        name: 'outlook_create_event',
        description: 'Create a new calendar event in Outlook',
        inputSchema: {
          type: 'object',
          properties: {
            subject: {
              type: 'string',
              description: 'Event subject/title',
            },
            start: {
              type: 'object',
              properties: {
                dateTime: {
                  type: 'string',
                  description: 'Start date/time in ISO 8601 format',
                },
                timeZone: {
                  type: 'string',
                  description: 'Time zone (e.g., "Pacific Standard Time")',
                },
              },
              required: ['dateTime', 'timeZone'],
            },
            end: {
              type: 'object',
              properties: {
                dateTime: {
                  type: 'string',
                  description: 'End date/time in ISO 8601 format',
                },
                timeZone: {
                  type: 'string',
                  description: 'Time zone (e.g., "Pacific Standard Time")',
                },
              },
              required: ['dateTime', 'timeZone'],
            },
            body: {
              type: 'string',
              description: 'Event description',
            },
            location: {
              type: 'string',
              description: 'Event location',
            },
            attendees: {
              type: 'array',
              items: { type: 'string' },
              description: 'Attendee email addresses',
            },
          },
          required: ['subject', 'start', 'end'],
        },
      },
    ],
  };
});

server.setRequestHandler('tools/call', async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case 'outlook_authenticate':
        return await authenticateTool(authManager);
      
      case 'outlook_list_emails':
        return await listEmailsTool(authManager, args);
      
      case 'outlook_send_email':
        return await sendEmailTool(authManager, args);
      
      case 'outlook_list_events':
        return await listEventsTool(authManager, args);
      
      case 'outlook_create_event':
        return await createEventTool(authManager, args);
      
      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    return {
      error: {
        code: 'TOOL_ERROR',
        message: error.message,
      },
    };
  }
});

async function main() {
  if (!process.env.AZURE_CLIENT_ID || !process.env.AZURE_TENANT_ID) {
    console.error('Error: AZURE_CLIENT_ID and AZURE_TENANT_ID environment variables are required.');
    console.error('Please set these in your MCP server configuration.');
    process.exit(1);
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Outlook MCP server started');
}

main().catch((error) => {
  console.error('Server error:', error);
  process.exit(1);
});