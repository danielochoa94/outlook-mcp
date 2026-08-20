# M365 Assistant MCP Server

A local Model Context Protocol server for working with Outlook, Microsoft To Do, and SharePoint through Microsoft Graph.

The server runs over stdio and works with Claude Desktop and other MCP-compatible clients.

## Features

- Read, search, draft, schedule, send, organize, and delete Outlook email.
- List, create, cancel, decline, and delete calendar events.
- Create, move, rename, empty, and delete mail folders.
- List, create, reorder, and delete inbox rules.
- Manage Microsoft To Do lists and tasks.
- Browse SharePoint sites and document libraries, and download, upload, move, rename, or delete files.
- Refresh Microsoft Graph access tokens automatically.
- Run against mock data for development and testing.

## Tools

The server currently exposes 45 tools.

### Authentication

| Tool | Description |
|---|---|
| `about` | Describe the server. |
| `authenticate` | Start Microsoft Graph authentication. |
| `check-auth-status` | Check whether a valid access token is available. |

### Email

| Tool | Description |
|---|---|
| `list-emails` | List messages in a mail folder with pagination. |
| `search-emails` | Search messages by text, sender, recipient, subject, attachment status, or read status. |
| `read-email` | Read a message with sanitized plain-text content. |
| `send-email` | Send a plain-text or HTML message with optional CC, BCC, and importance. |
| `reply-email` | Reply (or reply-all) in the original conversation thread; send now, schedule, or save as a draft. |
| `forward-email` | Forward a message in its thread; send now, schedule, or save as a draft. |
| `schedule-email` | Schedule a message for future delivery by Exchange. |
| `list-scheduled-emails` | List messages awaiting a future delivery time. |
| `cancel-scheduled-email` | Cancel a scheduled message by moving it to Deleted Items. |
| `draft-email` | Save an email draft. |
| `mark-as-read` | Mark a message as read or unread. |
| `delete-email` | Move a message to Deleted Items or permanently delete it. |

### Calendar

| Tool | Description |
|---|---|
| `list-events` | List events in a date range. |
| `create-event` | Create an event with optional attendees and body content. |
| `decline-event` | Decline an invitation with an optional comment. |
| `cancel-event` | Cancel an event with an optional comment. |
| `delete-event` | Delete an event. |

### Mail folders

| Tool | Description |
|---|---|
| `list-folders` | List mail folders, item counts, and optional child folders. |
| `create-folder` | Create a root or child mail folder. |
| `delete-folder` | Delete an empty folder or recursively delete a folder and its contents. |
| `move-emails` | Move selected messages between folders. |
| `move-folder` | Move or rename a mail folder. |
| `move-folder-contents` | Move every message from one folder to another using pagination and batching. |

### Inbox rules

| Tool | Description |
|---|---|
| `list-rules` | List inbox rules with optional condition and action details. |
| `create-rule` | Create a rule that can move, mark, or forward matching messages. |
| `edit-rule-sequence` | Change a rule's execution order. |
| `delete-rule` | Delete a rule by name. |

### Microsoft To Do

| Tool | Description |
|---|---|
| `list-todo-lists` | List task lists. |
| `create-todo-list` | Create a task list. |
| `update-todo-list` | Rename a task list. |
| `delete-todo-list` | Delete a task list. |
| `list-tasks` | List all, active, or completed tasks in a list. |
| `create-task` | Create a task with optional notes, due date, and importance. |
| `update-task` | Update a task's details or status. |
| `complete-task` | Mark a task as completed. |
| `delete-task` | Delete a task. |
| `move-task` | Copy a task to another list and delete the original. |

### SharePoint

| Tool | Description |
|---|---|
| `list-sites` | List or search accessible SharePoint sites. |
| `list-document-libraries` | List document libraries in a site. |
| `list-sharepoint-files` | Browse files and folders in a document library. |
| `download-sharepoint-file` | Download a file to the local filesystem. |
| `upload-sharepoint-file` | Upload a local file, using a chunked upload for large files. |
| `delete-sharepoint-file` | Move a file or folder to the site recycle bin. |
| `move-sharepoint-file` | Move or rename a file or folder. |

SharePoint requires a Microsoft work or school account.
Some tenants require administrator consent for SharePoint permissions.

## Requirements

- Node.js 18 or later.
- npm.
- A Microsoft Entra app registration.
- A Microsoft 365 account with access to the services you intend to use.

## Installation

Install the dependencies:

```bash
npm install
```

## Microsoft Entra setup

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com/) and select **App registrations**.
2. Create an app registration for the Microsoft accounts you intend to use.
3. Add `http://localhost:3333/auth/callback` as a redirect URI.
4. Enable public client flows under **Authentication**.
5. Copy the application client ID.
6. Add the delegated Microsoft Graph permissions listed below.

The server uses authorization code flow with PKCE and does not require a client secret.

### Delegated permissions

- `offline_access`
- `User.Read`
- `Mail.Read`
- `Mail.ReadWrite`
- `Mail.Send`
- `MailboxSettings.ReadWrite`
- `Calendars.Read`
- `Calendars.ReadWrite`
- `Files.Read`
- `Files.ReadWrite`
- `Tasks.ReadWrite`
- `Sites.ReadWrite.All`

Grant administrator consent when required by your tenant.

## Configuration

The server reads configuration from environment variables and from a `.env` file in the repository root.

```dotenv
OUTLOOK_CLIENT_ID=your-application-client-id
MS_TENANT_ID=your-tenant-id
USE_TEST_MODE=false
```

`MS_TENANT_ID` defaults to `common`.
Use your directory tenant ID for a single-tenant app.

Sovereign cloud users can override the authority host:

```dotenv
MS_AUTHORITY_HOST=https://login.microsoftonline.us
```

## MCP client configuration

Use an absolute WSL path to `index.js`.

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": ["/home/your-user/projects/outlook-mcp/index.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-application-client-id",
        "MS_TENANT_ID": "your-tenant-id",
        "USE_TEST_MODE": "false"
      }
    }
  }
}
```

The MCP client must be able to launch the WSL `node` executable and access the configured path.

## Authentication

Start the local callback server in a terminal:

```bash
npm run auth-server
```

Then call the `authenticate` tool from your MCP client and open the returned URL.
After Microsoft redirects to `http://localhost:3333/auth/callback`, the server saves the tokens to `~/.outlook-mcp-tokens.json` with permissions `0600`.

The MCP server refreshes expired access tokens automatically.
The callback server does not need to remain running after authentication completes.

To sign in again, stop the MCP server, remove `~/.outlook-mcp-tokens.json`, restart the callback server, and authenticate again.

## Development

Run the Jest suite:

```bash
npm test
```

Run the server against mock Graph responses:

```bash
npm run test-mode
```

Open the MCP Inspector:

```bash
npm run inspect
```

Start the stdio server directly:

```bash
npm start
```

## Notes

- `download-sharepoint-file` and `upload-sharepoint-file` read or write paths on the machine running the MCP server.
- Permanent email deletion is irreversible.
- Deleting a SharePoint item moves it to the site's recycle bin.
- `move-task` is implemented as create-then-delete because Microsoft Graph has no native task-move operation.

## License

MIT
