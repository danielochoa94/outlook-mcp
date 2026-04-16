/**
 * Update To Do task list functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Update To Do task list handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateTodoList(args) {
  const { listId, displayName } = args;

  if (!listId) {
    return {
      content: [{
        type: "text",
        text: "List ID is required."
      }]
    };
  }

  if (!displayName) {
    return {
      content: [{
        type: "text",
        text: "Display name is required to update a task list."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/todo/lists/${listId}`;
    const body = { displayName };

    const response = await callGraphAPI(accessToken, 'PATCH', endpoint, body);

    return {
      content: [{
        type: "text",
        text: `Task list renamed to '${response.displayName}' successfully.`
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
        text: `Error updating task list: ${error.message}`
      }]
    };
  }
}

module.exports = handleUpdateTodoList;
