/**
 * Delete folder functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getFolderIdByName } = require('../email/folder-utils');

/**
 * Delete folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteFolder(args) {
  const folderPath = args.folder;
  const recursive = args.recursive || false;
  const force = args.force || false;

  if (!folderPath) {
    return {
      content: [{
        type: "text",
        text: "Folder name or path is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const folderId = await getFolderIdByName(accessToken, folderPath);
    if (!folderId) {
      return {
        content: [{
          type: "text",
          text: `Folder "${folderPath}" not found.`
        }]
      };
    }

    // Check folder contents before deleting
    const folderInfo = await callGraphAPI(
      accessToken,
      'GET',
      `me/mailFolders/${folderId}`,
      null,
      { $select: 'id,displayName,totalItemCount,childFolderCount' }
    );

    if (!force) {
      if (folderInfo.totalItemCount > 0) {
        return {
          content: [{
            type: "text",
            text: `Folder "${folderPath}" still contains ${folderInfo.totalItemCount} item(s). Use force=true to delete anyway (contents go to Deleted Items), or move them first.`
          }]
        };
      }

      if (recursive && folderInfo.childFolderCount > 0) {
        const nonEmptyChildren = await findNonEmptyChildren(accessToken, folderId);
        if (nonEmptyChildren.length > 0) {
          const names = nonEmptyChildren.map(folder => `"${folder.displayName}" (${folder.totalItemCount} items)`).join(', ');
          return {
            content: [{
              type: "text",
              text: `These subfolders still have items: ${names}. Use force=true to delete anyway (contents go to Deleted Items), or move them first.`
            }]
          };
        }
      }
    }

    if (folderInfo.childFolderCount > 0 && !recursive) {
      return {
        content: [{
          type: "text",
          text: `Folder "${folderPath}" has ${folderInfo.childFolderCount} child folder(s). Set recursive=true to delete it and all subfolders.`
        }]
      };
    }

    // Delete the folder (Graph API deletes children recursively)
    await callGraphAPI(accessToken, 'DELETE', `me/mailFolders/${folderId}`);

    const childNote = folderInfo.childFolderCount > 0
      ? ` and its ${folderInfo.childFolderCount} subfolder(s)`
      : '';

    return {
      content: [{
        type: "text",
        text: `Deleted folder "${folderPath}"${childNote}.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error deleting folder: ${error.message}`
      }]
    };
  }
}

/**
 * Recursively find child folders that still contain items
 * @param {string} accessToken - Access token
 * @param {string} folderId - Parent folder ID
 * @returns {Promise<Array>} - Array of non-empty folder objects
 */
async function findNonEmptyChildren(accessToken, folderId) {
  const nonEmpty = [];

  const response = await callGraphAPI(
    accessToken,
    'GET',
    `me/mailFolders/${folderId}/childFolders`,
    null,
    { $top: 100, $select: 'id,displayName,totalItemCount,childFolderCount' }
  );

  const children = response.value || [];

  for (const child of children) {
    if (child.totalItemCount > 0) {
      nonEmpty.push(child);
    }
    if (child.childFolderCount > 0) {
      const deepNonEmpty = await findNonEmptyChildren(accessToken, child.id);
      nonEmpty.push(...deepNonEmpty);
    }
  }

  return nonEmpty;
}

module.exports = handleDeleteFolder;
