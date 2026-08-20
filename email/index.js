/**
 * Email module for Outlook MCP server
 */
const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleDraftEmail = require('./draft');
const { handleReplyEmail, handleForwardEmail } = require('./reply');
const handleMarkAsRead = require('./mark-as-read');
const handleDeleteEmail = require('./delete');
const {
  handleScheduleEmail,
  handleListScheduledEmails,
  handleCancelScheduledEmail
} = require('./schedule');

// Email tool definitions
const emailTools = [
  {
    name: "list-emails",
    description: "Lists recent emails from your inbox",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"
        },
        count: {
          type: "number",
          description: "Number of emails to retrieve (default: 10, max: 50 per page)"
        },
        skip: {
          type: "number",
          description: "Number of emails to skip for pagination (default: 0). Use with count to page through results."
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text to find in emails"
        },
        folder: {
          type: "string",
          description: "Email folder to search in. Omit to search across all folders."
        },
        from: {
          type: "string",
          description: "Filter by sender email address or name"
        },
        to: {
          type: "string",
          description: "Filter by recipient email address or name"
        },
        subject: {
          type: "string",
          description: "Filter by email subject"
        },
        hasAttachments: {
          type: "boolean",
          description: "Filter to only emails with attachments"
        },
        unreadOnly: {
          type: "boolean",
          description: "Filter to only unread emails"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50 per page)"
        },
        skip: {
          type: "number",
          description: "Number of emails to skip for pagination (default: 0). Use with count to page through results."
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email. HTML emails are sanitized to plain text.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to read"
        },
        maxChars: {
          type: "number",
          description: "Maximum characters of body to return (default: 5000, use 0 for unlimited)"
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Composes and sends a new email. Supports both plain text and HTML content.",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (plain text or HTML)"
        },
        isHtml: {
          type: "boolean",
          description: "Set to true to send as HTML, false for plain text. If not specified, auto-detects based on <html> tag presence."
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        saveToSentItems: {
          type: "boolean",
          description: "Whether to save the email to sent items"
        }
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "reply-email",
    description: "Replies to an existing email, keeping it in the same Outlook conversation thread. Sends immediately by default; use sendAt to schedule it or saveAsDraft to leave it in Drafts.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to reply to"
        },
        body: {
          type: "string",
          description: "Your reply text, placed above the quoted original"
        },
        replyAll: {
          type: "boolean",
          description: "Reply to every recipient of the original instead of just the sender. Default: false"
        },
        to: {
          type: "string",
          description: "Comma-separated addresses to add to the recipients Outlook already fills in"
        },
        cc: {
          type: "string",
          description: "Comma-separated addresses to add to CC"
        },
        bcc: {
          type: "string",
          description: "Comma-separated addresses to add to BCC"
        },
        isHtml: {
          type: "boolean",
          description: "Set to true to treat body as HTML, false for plain text. If not specified, auto-detects based on <html> tag presence."
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        sendAt: {
          type: "string",
          description: "Schedule the reply instead of sending now, as ISO 8601 with an explicit timezone (e.g. '2025-01-31T09:00:00-05:00'). Must be in the future. Cancel with 'cancel-scheduled-email'."
        },
        saveAsDraft: {
          type: "boolean",
          description: "Save the reply to Drafts instead of sending it. Cannot be combined with sendAt. Default: false"
        }
      },
      required: ["id", "body"]
    },
    handler: handleReplyEmail
  },
  {
    name: "forward-email",
    description: "Forwards an existing email, keeping it in the same Outlook conversation thread. Sends immediately by default; use sendAt to schedule it or saveAsDraft to leave it in Drafts.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to forward"
        },
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        body: {
          type: "string",
          description: "Your comment, placed above the forwarded message"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        isHtml: {
          type: "boolean",
          description: "Set to true to treat body as HTML, false for plain text. If not specified, auto-detects based on <html> tag presence."
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        sendAt: {
          type: "string",
          description: "Schedule the forward instead of sending now, as ISO 8601 with an explicit timezone (e.g. '2025-01-31T09:00:00-05:00'). Must be in the future. Cancel with 'cancel-scheduled-email'."
        },
        saveAsDraft: {
          type: "boolean",
          description: "Save the forward to Drafts instead of sending it. Cannot be combined with sendAt. Default: false"
        }
      },
      required: ["id", "to"]
    },
    handler: handleForwardEmail
  },
  {
    name: "schedule-email",
    description: "Composes an email and schedules it to be sent at a future time. The message waits in Drafts and is delivered by Exchange at the requested time, with no client running.",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (plain text or HTML)"
        },
        sendAt: {
          type: "string",
          description: "When to send, as ISO 8601 with an explicit timezone (e.g. '2025-01-31T09:00:00-05:00' or '2025-01-31T14:00:00Z'). Must be in the future."
        },
        isHtml: {
          type: "boolean",
          description: "Set to true to send as HTML, false for plain text. If not specified, auto-detects based on <html> tag presence."
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        }
      },
      required: ["to", "subject", "body", "sendAt"]
    },
    handler: handleScheduleEmail
  },
  {
    name: "list-scheduled-emails",
    description: "Lists emails that are queued for a future send time, with their IDs and send times",
    inputSchema: {
      type: "object",
      properties: {
        count: {
          type: "number",
          description: "Maximum number of messages to inspect per folder (default: 25, max: 50)"
        }
      },
      required: []
    },
    handler: handleListScheduledEmails
  },
  {
    name: "cancel-scheduled-email",
    description: "Cancels a scheduled email before it is sent by deleting the queued message (moves it to Deleted Items)",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the scheduled message, from 'list-scheduled-emails'"
        }
      },
      required: ["id"]
    },
    handler: handleCancelScheduledEmail
  },
  {
    name: "draft-email",
    description: "Creates and saves an email draft in Outlook",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Draft email subject"
        },
        body: {
          type: "string",
          description: "Draft email body content (can be plain text or HTML)"
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        }
      },
      required: []
    },
    handler: handleDraftEmail
  },
  {
    name: "mark-as-read",
    description: "Marks an email as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to mark as read/unread"
        },
        isRead: {
          type: "boolean",
          description: "Whether to mark as read (true) or unread (false). Default: true"
        }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  },
  {
    name: "delete-email",
    description: "Deletes an email by moving it to Deleted Items (trash). Use permanent=true to hard delete.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to delete"
        },
        permanent: {
          type: "boolean",
          description: "If true, permanently delete the email instead of moving to Deleted Items. Default: false"
        }
      },
      required: ["id"]
    },
    handler: handleDeleteEmail
  }
];

module.exports = {
  emailTools,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleReplyEmail,
  handleForwardEmail,
  handleDraftEmail,
  handleScheduleEmail,
  handleListScheduledEmails,
  handleCancelScheduledEmail,
  handleMarkAsRead,
  handleDeleteEmail
};
