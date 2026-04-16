/**
 * SharePoint module for M365 Assistant MCP server
 */
const handleListSites = require('./list-sites');
const handleListDocumentLibraries = require('./list-document-libraries');
const handleListFiles = require('./list-files');
const handleDownloadFile = require('./download-file');
const handleUploadFile = require('./upload-file');
const handleDeleteFile = require('./delete-file');
const handleMoveFile = require('./move-file');

// SharePoint tool definitions
const sharepointTools = [
  {
    name: "list-sites",
    description: "Lists or searches SharePoint sites you have access to. Use the search parameter to find sites by name.",
    inputSchema: {
      type: "object",
      properties: {
        search: {
          type: "string",
          description: "Search keyword to find sites (omit to list all accessible sites)"
        },
        count: {
          type: "number",
          description: "Number of sites to retrieve (default: 25, max: 50)"
        }
      },
      required: []
    },
    handler: handleListSites
  },
  {
    name: "list-document-libraries",
    description: "Lists document libraries (drives) in a SharePoint site",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        }
      },
      required: ["siteId"]
    },
    handler: handleListDocumentLibraries
  },
  {
    name: "list-sharepoint-files",
    description: "Lists files and folders in a SharePoint document library. Provide a folderId to browse into a subfolder.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        driveId: {
          type: "string",
          description: "The ID of the document library (drive)"
        },
        folderId: {
          type: "string",
          description: "Optional folder ID to list contents of a specific folder"
        },
        count: {
          type: "number",
          description: "Number of items to retrieve (default: 25, max: 50)"
        }
      },
      required: ["siteId", "driveId"]
    },
    handler: handleListFiles
  },
  {
    name: "download-sharepoint-file",
    description: "Downloads a file from a SharePoint document library to a local path",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        driveId: {
          type: "string",
          description: "The ID of the document library (drive)"
        },
        itemId: {
          type: "string",
          description: "The ID of the file to download"
        },
        savePath: {
          type: "string",
          description: "Local file path or directory to save the downloaded file"
        }
      },
      required: ["siteId", "driveId", "itemId", "savePath"]
    },
    handler: handleDownloadFile
  },
  {
    name: "upload-sharepoint-file",
    description: "Uploads a local file to a SharePoint document library. Supports files up to 250MB with chunked upload for large files.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        driveId: {
          type: "string",
          description: "The ID of the document library (drive)"
        },
        localPath: {
          type: "string",
          description: "Local file path to upload"
        },
        destinationPath: {
          type: "string",
          description: "Optional folder path in SharePoint (e.g., 'Reports/2026')"
        },
        folderId: {
          type: "string",
          description: "Optional folder ID to upload into (alternative to destinationPath)"
        }
      },
      required: ["siteId", "driveId", "localPath"]
    },
    handler: handleUploadFile
  },
  {
    name: "delete-sharepoint-file",
    description: "Deletes a file or folder from a SharePoint document library. The item is moved to the site recycle bin.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        driveId: {
          type: "string",
          description: "The ID of the document library (drive)"
        },
        itemId: {
          type: "string",
          description: "The ID of the file or folder to delete"
        }
      },
      required: ["siteId", "driveId", "itemId"]
    },
    handler: handleDeleteFile
  },
  {
    name: "move-sharepoint-file",
    description: "Moves and/or renames a file or folder in a SharePoint document library. Provide newName to rename, destinationFolderId to move, or both.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        driveId: {
          type: "string",
          description: "The ID of the document library (drive)"
        },
        itemId: {
          type: "string",
          description: "The ID of the file or folder to move/rename"
        },
        newName: {
          type: "string",
          description: "New name for the file or folder (include file extension)"
        },
        destinationFolderId: {
          type: "string",
          description: "ID of the destination folder to move the item into"
        }
      },
      required: ["siteId", "driveId", "itemId"]
    },
    handler: handleMoveFile
  }
];

module.exports = {
  sharepointTools,
  handleListSites,
  handleListDocumentLibraries,
  handleListFiles,
  handleDownloadFile,
  handleUploadFile,
  handleDeleteFile,
  handleMoveFile
};
