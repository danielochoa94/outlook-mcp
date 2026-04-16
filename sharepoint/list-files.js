/**
 * List files and folders in a SharePoint document library
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List files/folders in a SharePoint document library
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFiles(args) {
  const { siteId, driveId, folderId, count } = args;

  if (!siteId || !driveId) {
    return {
      content: [{
        type: "text",
        text: "Site ID and Drive ID are required. Use 'list-sites' and 'list-document-libraries' to find them."
      }]
    };
  }

  const maxCount = Math.min(count || 25, config.MAX_RESULT_COUNT);

  try {
    const accessToken = await ensureAuthenticated();

    let endpoint;
    if (folderId) {
      endpoint = `sites/${siteId}/drives/${driveId}/items/${folderId}/children`;
    } else {
      endpoint = `sites/${siteId}/drives/${driveId}/root/children`;
    }

    const queryParams = {
      $top: maxCount,
      $select: config.SHAREPOINT_ITEM_SELECT_FIELDS,
      $orderby: 'name'
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No files or folders found."
        }]
      };
    }

    const itemOutput = response.value.map((item, index) => {
      const isFolder = !!item.folder;
      const icon = isFolder ? '[Folder]' : '[File]';
      const size = !isFolder && item.size ? `\n   Size: ${formatSize(item.size)}` : '';
      const childCount = isFolder && item.folder.childCount !== undefined
        ? `\n   Items: ${item.folder.childCount}`
        : '';
      const modified = item.lastModifiedDateTime
        ? `\n   Modified: ${item.lastModifiedDateTime.split('T')[0]}`
        : '';
      return `${index + 1}. ${icon} ${item.name}${size}${childCount}${modified}\n   ID: ${item.id}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} item(s):\n\n${itemOutput}`
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
        text: `Error listing files: ${error.message}`
      }]
    };
  }
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

module.exports = handleListFiles;
