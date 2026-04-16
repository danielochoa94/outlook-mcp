/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');
require('dotenv').config({ path: path.join(__dirname, '.env') });

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

module.exports = {
  // Server information
  SERVER_NAME: "m365-assistant",
  SERVER_VERSION: "2.0.0",
  
  // Test mode setting
  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',
  
  // Authentication configuration
  AUTH_CONFIG: {
    clientId: process.env.OUTLOOK_CLIENT_ID || '',
    redirectUri: 'http://localhost:3333/auth/callback',
    scopes: ['Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'MailboxSettings.ReadWrite', 'User.Read', 'Calendars.Read', 'Calendars.ReadWrite', 'Files.Read', 'Files.ReadWrite', 'Tasks.ReadWrite', 'Sites.ReadWrite.All'],
    tokenStorePath: path.join(homeDir, '.outlook-mcp-tokens.json'),
    authServerUrl: 'http://localhost:3333'
  },
  
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,start,end,location,bodyPreview,isAllDay,recurrence,attendees',

  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,

  // Timezone
  DEFAULT_TIMEZONE: "Central European Standard Time",

  // OneDrive constants
  ONEDRIVE_SELECT_FIELDS: 'id,name,size,lastModifiedDateTime,webUrl,folder,file,parentReference',
  ONEDRIVE_UPLOAD_THRESHOLD: 4 * 1024 * 1024, // 4MB - files larger than this need chunked upload

  // To Do constants
  TODO_LIST_SELECT_FIELDS: 'id,displayName,isOwner,isShared,wellknownListName',
  TODO_TASK_SELECT_FIELDS: 'id,title,body,status,importance,createdDateTime,lastModifiedDateTime,dueDateTime,completedDateTime,linkedResources',

  // SharePoint constants
  SHAREPOINT_SITE_SELECT_FIELDS: 'id,displayName,description,webUrl,createdDateTime,lastModifiedDateTime',
  SHAREPOINT_DRIVE_SELECT_FIELDS: 'id,name,description,webUrl,driveType,createdDateTime,lastModifiedDateTime',
  SHAREPOINT_ITEM_SELECT_FIELDS: 'id,name,size,lastModifiedDateTime,webUrl,folder,file,parentReference,createdBy,lastModifiedBy',
};
