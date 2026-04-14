/**
 * Folder management module for Outlook MCP server
 */
const handleListFolders = require('./list');
const handleCreateFolder = require('./create');
const handleMoveEmails = require('./move');
const handleMoveFolder = require('./move-folder');

// Folder management tool definitions
const folderTools = [
  {
    name: "list-folders",
    description: "Lists mail folders in your Outlook account",
    inputSchema: {
      type: "object",
      properties: {
        includeItemCounts: {
          type: "boolean",
          description: "Include counts of total and unread items"
        },
        includeChildren: {
          type: "boolean",
          description: "Include child folders in hierarchy"
        }
      },
      required: []
    },
    handler: handleListFolders
  },
  {
    name: "create-folder",
    description: "Creates a new mail folder",
    inputSchema: {
      type: "object",
      properties: {
        name: {
          type: "string",
          description: "Name of the folder to create"
        },
        parentFolder: {
          type: "string",
          description: "Optional parent folder name (default is root)"
        }
      },
      required: ["name"]
    },
    handler: handleCreateFolder
  },
  {
    name: "move-emails",
    description: "Moves emails from one folder to another",
    inputSchema: {
      type: "object",
      properties: {
        emailIds: {
          type: "string",
          description: "Comma-separated list of email IDs to move"
        },
        targetFolder: {
          type: "string",
          description: "Name of the folder to move emails to"
        },
        sourceFolder: {
          type: "string",
          description: "Optional name of the source folder (default is inbox)"
        }
      },
      required: ["emailIds", "targetFolder"]
    },
    handler: handleMoveEmails
  },
  {
    name: "move-folder",
    description: "Moves a mail folder to a new parent folder and/or renames it",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Name or path of the folder to move/rename (e.g. 'Projects' or 'Work/Projects')"
        },
        destinationFolder: {
          type: "string",
          description: "Name or path of the new parent folder to move it into"
        },
        newName: {
          type: "string",
          description: "New display name for the folder"
        }
      },
      required: ["folder"]
    },
    handler: handleMoveFolder
  }
];

module.exports = {
  folderTools,
  handleListFolders,
  handleCreateFolder,
  handleMoveEmails,
  handleMoveFolder
};
