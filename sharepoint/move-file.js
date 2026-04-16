/**
 * Move or rename a file/folder in SharePoint
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Move and/or rename a file or folder in a SharePoint document library
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveFile(args) {
  const { siteId, driveId, itemId, newName, destinationFolderId } = args;

  if (!siteId || !driveId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "Site ID, Drive ID, and Item ID are required. Use 'list-sharepoint-files' to find them."
      }]
    };
  }

  if (!newName && !destinationFolderId) {
    return {
      content: [{
        type: "text",
        text: "At least one of newName or destinationFolderId must be provided."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Get current item metadata
    const metadata = await callGraphAPI(
      accessToken,
      'GET',
      `sites/${siteId}/drives/${driveId}/items/${itemId}`,
      null,
      { $select: 'name,parentReference,folder,file' }
    );

    const itemType = metadata.folder ? 'folder' : 'file';
    const originalName = metadata.name;

    // Build PATCH body
    const patchBody = {};

    if (newName) {
      patchBody.name = newName;
    }

    if (destinationFolderId) {
      patchBody.parentReference = { id: destinationFolderId };
    }

    // Apply move/rename via PATCH
    const result = await callGraphAPI(
      accessToken,
      'PATCH',
      `sites/${siteId}/drives/${driveId}/items/${itemId}`,
      patchBody
    );

    // Build response message
    const actions = [];
    if (newName && newName !== originalName) {
      actions.push(`renamed from "${originalName}" to "${newName}"`);
    }
    if (destinationFolderId) {
      actions.push(`moved to folder "${result.parentReference?.name || destinationFolderId}"`);
    }

    const actionText = actions.length > 0
      ? actions.join(' and ')
      : 'updated (no changes detected)';

    return {
      content: [{
        type: "text",
        text: `${itemType.charAt(0).toUpperCase() + itemType.slice(1)} ${actionText}.\n   Name: ${result.name}\n   URL: ${result.webUrl}\n   ID: ${result.id}`
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
        text: `Error moving/renaming item: ${error.message}`
      }]
    };
  }
}

module.exports = handleMoveFile;
