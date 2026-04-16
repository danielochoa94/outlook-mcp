/**
 * Delete a file or folder from SharePoint
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Delete a file or folder from a SharePoint document library
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteFile(args) {
  const { siteId, driveId, itemId } = args;

  if (!siteId || !driveId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "Site ID, Drive ID, and Item ID are required. Use 'list-sharepoint-files' to find them."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Get item metadata first to confirm what we're deleting
    const metadata = await callGraphAPI(
      accessToken,
      'GET',
      `sites/${siteId}/drives/${driveId}/items/${itemId}`,
      null,
      { $select: 'name,size,file,folder' }
    );

    const itemType = metadata.folder ? 'folder' : 'file';
    const itemName = metadata.name;

    // Delete the item
    await callGraphAPI(
      accessToken,
      'DELETE',
      `sites/${siteId}/drives/${driveId}/items/${itemId}`
    );

    return {
      content: [{
        type: "text",
        text: `Deleted ${itemType} "${itemName}" successfully. It has been moved to the site recycle bin.`
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
        text: `Error deleting item: ${error.message}`
      }]
    };
  }
}

module.exports = handleDeleteFile;
