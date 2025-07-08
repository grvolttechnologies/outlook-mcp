#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { OutlookAuthManager } from './auth/auth.js';
import { 
  authenticateTool,
  listEmailsTool,
  sendEmailTool,
  listEventsTool,
  createEventTool,
  getEmailTool,
  searchEmailsTool,
  createDraftTool,
  replyToEmailTool,
  replyAllTool,
  forwardEmailTool,
  deleteEmailTool
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
  process.env.AZURE_TENANT_ID,
  process.env.AZURE_CLIENT_SECRET
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
      {
        name: 'outlook_get_email',
        description: 'Get detailed information about a specific email',
        inputSchema: {
          type: 'object',
          properties: {
            messageId: {
              type: 'string',
              description: 'The ID of the email message to retrieve',
            },
          },
          required: ['messageId'],
        },
      },
      {
        name: 'outlook_search_emails',
        description: 'Search emails across all folders with advanced filters for analysis',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Free-text search query across email content',
            },
            subject: {
              type: 'string',
              description: 'Search emails with specific subject text',
            },
            from: {
              type: 'string',
              description: 'Filter emails from specific sender',
            },
            startDate: {
              type: 'string',
              description: 'Start date for email search (ISO 8601 format)',
            },
            endDate: {
              type: 'string',
              description: 'End date for email search (ISO 8601 format)',
            },
            folders: {
              type: 'array',
              items: { type: 'string' },
              description: 'Specific folders to search in',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of emails to return (max 1000)',
              default: 100,
            },
            includeBody: {
              type: 'boolean',
              description: 'Include full email body content for analysis',
              default: true,
            },
            orderBy: {
              type: 'string',
              description: 'Sort order (e.g., "receivedDateTime desc")',
              default: 'receivedDateTime desc',
            },
          },
        },
      },
      {
        name: 'outlook_create_draft',
        description: 'Create an email draft without sending',
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
            importance: {
              type: 'string',
              enum: ['low', 'normal', 'high'],
              default: 'normal',
              description: 'Email importance level',
            },
          },
          required: ['to', 'subject'],
        },
      },
      {
        name: 'outlook_reply_to_email',
        description: 'Reply to an existing email',
        inputSchema: {
          type: 'object',
          properties: {
            messageId: {
              type: 'string',
              description: 'The ID of the email to reply to',
            },
            body: {
              type: 'string',
              description: 'Reply message body',
            },
            bodyType: {
              type: 'string',
              enum: ['text', 'html'],
              default: 'text',
              description: 'Body content type',
            },
            comment: {
              type: 'string',
              description: 'Optional comment for the reply',
            },
          },
          required: ['messageId'],
        },
      },
      {
        name: 'outlook_reply_all',
        description: 'Reply to all recipients of an existing email',
        inputSchema: {
          type: 'object',
          properties: {
            messageId: {
              type: 'string',
              description: 'The ID of the email to reply all to',
            },
            body: {
              type: 'string',
              description: 'Reply message body',
            },
            bodyType: {
              type: 'string',
              enum: ['text', 'html'],
              default: 'text',
              description: 'Body content type',
            },
            comment: {
              type: 'string',
              description: 'Optional comment for the reply',
            },
          },
          required: ['messageId'],
        },
      },
      {
        name: 'outlook_forward_email',
        description: 'Forward an existing email to new recipients',
        inputSchema: {
          type: 'object',
          properties: {
            messageId: {
              type: 'string',
              description: 'The ID of the email to forward',
            },
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Recipient email addresses for forwarding',
            },
            body: {
              type: 'string',
              description: 'Additional message body for the forward',
            },
            bodyType: {
              type: 'string',
              enum: ['text', 'html'],
              default: 'text',
              description: 'Body content type',
            },
            comment: {
              type: 'string',
              description: 'Optional comment for the forward',
            },
          },
          required: ['messageId', 'to'],
        },
      },
      {
        name: 'outlook_delete_email',
        description: 'Delete an email (move to Deleted Items or permanently delete)',
        inputSchema: {
          type: 'object',
          properties: {
            messageId: {
              type: 'string',
              description: 'The ID of the email to delete',
            },
            permanentDelete: {
              type: 'boolean',
              description: 'Whether to permanently delete (true) or move to Deleted Items (false)',
              default: false,
            },
          },
          required: ['messageId'],
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
      
      case 'outlook_get_email':
        return await getEmailTool(authManager, args);
      
      case 'outlook_search_emails':
        return await searchEmailsTool(authManager, args);
      
      case 'outlook_create_draft':
        return await createDraftTool(authManager, args);
      
      case 'outlook_reply_to_email':
        return await replyToEmailTool(authManager, args);
      
      case 'outlook_reply_all':
        return await replyAllTool(authManager, args);
      
      case 'outlook_forward_email':
        return await forwardEmailTool(authManager, args);
      
      case 'outlook_delete_email':
        return await deleteEmailTool(authManager, args);
      
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
    console.error('Note: AZURE_CLIENT_SECRET is optional but recommended for server applications.');
    process.exit(1);
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Outlook MCP server started with secure token management');
}

main().catch((error) => {
  console.error('Server error:', error);
  process.exit(1);
});