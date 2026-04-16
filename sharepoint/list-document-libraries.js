/**
 * List document libraries in a SharePoint site
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List document libraries in a SharePoint site
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListDocumentLibraries(args) {
  const { siteId } = args;

  if (!siteId) {
    return {
      content: [{
        type: "text",
        text: "Site ID is required. Use 'list-sites' to find available sites."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `sites/${siteId}/drives`;
    const queryParams = {
      $select: config.SHAREPOINT_DRIVE_SELECT_FIELDS
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No document libraries found in this site."
        }]
      };
    }

    const libraryOutput = response.value.map((drive, index) => {
      const description = drive.description ? `\n   Description: ${drive.description}` : '';
      const type = drive.driveType ? `\n   Type: ${drive.driveType}` : '';
      return `${index + 1}. ${drive.name}${description}${type}\n   ID: ${drive.id}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} document library/libraries:\n\n${libraryOutput}`
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
        text: `Error listing document libraries: ${error.message}`
      }]
    };
  }
}

module.exports = handleListDocumentLibraries;
