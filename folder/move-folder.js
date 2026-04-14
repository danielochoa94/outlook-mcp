/**
 * Move/rename folder functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getFolderIdByName } = require('../email/folder-utils');

/**
 * Move folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveFolder(args) {
  const folderName = args.folder || '';
  const destinationFolder = args.destinationFolder || '';
  const newName = args.newName || '';

  if (!folderName) {
    return {
      content: [{
        type: "text",
        text: "Folder name is required. Specify the folder to move or rename."
      }]
    };
  }

  if (!destinationFolder && !newName) {
    return {
      content: [{
        type: "text",
        text: "At least one of destinationFolder or newName must be provided."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Resolve the source folder
    const folderId = await getFolderIdByName(accessToken, folderName);
    if (!folderId) {
      return {
        content: [{
          type: "text",
          text: `Folder "${folderName}" not found. Please specify a valid folder name or path.`
        }]
      };
    }

    // Build the patch body
    const patchBody = {};
    const actions = [];

    if (destinationFolder) {
      const destinationFolderId = await getFolderIdByName(accessToken, destinationFolder);
      if (!destinationFolderId) {
        return {
          content: [{
            type: "text",
            text: `Destination folder "${destinationFolder}" not found. Please specify a valid folder name or path.`
          }]
        };
      }
      patchBody.parentFolderId = destinationFolderId;
      actions.push(`moved to "${destinationFolder}"`);
    }

    if (newName) {
      patchBody.displayName = newName;
      actions.push(`renamed to "${newName}"`);
    }

    await callGraphAPI(
      accessToken,
      'PATCH',
      `me/mailFolders/${folderId}`,
      patchBody
    );

    return {
      content: [{
        type: "text",
        text: `Successfully ${actions.join(' and ')}: "${folderName}".`
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
        text: `Error moving/renaming folder: ${error.message}`
      }]
    };
  }
}

module.exports = handleMoveFolder;
