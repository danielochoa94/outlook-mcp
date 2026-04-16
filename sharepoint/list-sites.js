/**
 * List/search SharePoint sites
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List or search SharePoint sites
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListSites(args) {
  const { search, count } = args;
  const maxCount = Math.min(count || 25, config.MAX_RESULT_COUNT);

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = 'sites';
    const queryParams = {
      search: search || '*',
      $top: maxCount,
      $select: config.SHAREPOINT_SITE_SELECT_FIELDS
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: search
            ? `No SharePoint sites found matching "${search}".`
            : "No SharePoint sites found."
        }]
      };
    }

    const siteOutput = response.value.map((site, index) => {
      const description = site.description ? `\n   Description: ${site.description}` : '';
      const url = site.webUrl ? `\n   URL: ${site.webUrl}` : '';
      return `${index + 1}. ${site.displayName}${description}${url}\n   ID: ${site.id}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} SharePoint site(s):\n\n${siteOutput}`
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
        text: `Error listing sites: ${error.message}`
      }]
    };
  }
}

module.exports = handleListSites;
