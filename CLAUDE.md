# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

- `npm install` - **ALWAYS run first** to install dependencies
- `npm start` - Start the MCP server
- `npm run auth-server` - Start the OAuth authentication server on port 3333 (**required for authentication**)
- `npm run test-mode` - Start the server in test mode with mock data
- `npm run inspect` - Use MCP Inspector to test the server interactively
- `npm test` - Run Jest tests
- `npx kill-port 3333` - Kill process using port 3333 if auth server won't start

## Architecture Overview

This is a modular MCP (Model Context Protocol) server that provides Claude with access to Microsoft 365 services:
- **Outlook** - Email, calendar, folders, rules
- **To Do** - Task lists, tasks (GTD workflows)
- **SharePoint** - Sites, document libraries, files, lists

### Core Structure
- `index.js` - Main entry point that combines all module tools and handles MCP protocol
- `config.js` - Centralized configuration (API endpoints, scopes, field selections)
- `outlook-auth-server.js` - Standalone OAuth server for authentication flow

### Modules
Each module exports tools and handlers:
- `auth/` - OAuth 2.0 authentication with token management
- `calendar/` - Calendar operations (list, create, accept, decline, delete events)
- `email/` - Email management (list, search, read, send, mark as read)
- `folder/` - Folder operations (list, create, move)
- `rules/` - Email rules management
- `todo/` - Microsoft To Do operations (list/create/delete lists, list/create/update/complete/delete tasks)
- `sharepoint/` - SharePoint operations (list sites, document libraries, files, download, upload, lists, list items)
- `utils/` - Shared utilities including Graph API client and OData helpers

### Key Components
- **Token Management**: Tokens stored in `~/.outlook-mcp-tokens.json`
- **Graph API Client**: `utils/graph-api.js` handles Microsoft Graph API calls
- **Test Mode**: Mock data responses when `USE_TEST_MODE=true`
- **Modular Tools**: Each module exports tools array that gets combined in main server

## Authentication

### Graph API
1. Azure app registration required with permissions:
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Files.Read`, `Files.ReadWrite`
   - `Tasks.ReadWrite`
   - `Sites.ReadWrite.All`
   - `User.Read`, `offline_access`
2. Start auth server: `npm run auth-server`
3. Use authenticate tool to get OAuth URL
4. Complete browser authentication
5. Tokens automatically stored and refreshed

## Configuration

### Environment Variables
- **For .env file**: Use `MS_CLIENT_ID` and `MS_CLIENT_SECRET`
- **For Claude Desktop config**: Use `OUTLOOK_CLIENT_ID` and `OUTLOOK_CLIENT_SECRET`
- **Important**: Always use the client secret VALUE from Azure, not the Secret ID

### Config Constants
- `GRAPH_API_ENDPOINT`: `https://graph.microsoft.com/v1.0/`
- `ONEDRIVE_UPLOAD_THRESHOLD`: 4MB (files larger need chunked upload for SharePoint/OneDrive)
- Default page size: 25, max results: 50

### Common Setup Issues
1. **Missing dependencies**: Always run `npm install` first
2. **Wrong secret**: Use Azure secret VALUE, not ID (AADSTS7000215 error)
3. **Auth server not running**: Start `npm run auth-server` before authenticating
4. **Port conflicts**: Use `npx kill-port 3333` if port is in use

## Test Mode

Set `USE_TEST_MODE=true` to use mock data instead of real API calls. Mock responses defined in:
- `utils/mock-data.js` - Graph API mocks

## Error Handling

- Graph API auth failures: "UNAUTHORIZED" error
- API errors include status codes and response details
- Token expiration triggers re-authentication flow
- Empty API responses handled gracefully
