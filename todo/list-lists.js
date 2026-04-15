/**
 * List To Do task lists functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List To Do task lists handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListTodoLists(args) {
  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = 'me/todo/lists';

    const response = await callGraphAPI(accessToken, 'GET', endpoint);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No task lists found."
        }]
      };
    }

    const listOutput = response.value.map((list, index) => {
      const shared = list.isShared ? ' (shared)' : '';
      const wellknown = list.wellknownListName && list.wellknownListName !== 'none'
        ? ` [${list.wellknownListName}]`
        : '';
      return `${index + 1}. ${list.displayName}${wellknown}${shared}\n   ID: ${list.id}`;
    }).join("\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} task lists:\n\n${listOutput}`
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
        text: `Error listing task lists: ${error.message}`
      }]
    };
  }
}

module.exports = handleListTodoLists;