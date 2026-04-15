/**
 * Delete To Do task list functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Delete To Do task list handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteTodoList(args) {
  const { listId } = args;

  if (!listId) {
    return {
      content: [{
        type: "text",
        text: "List ID is required to delete a task list."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/todo/lists/${listId}`;

    await callGraphAPI(accessToken, 'DELETE', endpoint);

    return {
      content: [{
        type: "text",
        text: `Task list with ID ${listId} has been successfully deleted.`
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
        text: `Error deleting task list: ${error.message}`
      }]
    };
  }
}

module.exports = handleDeleteTodoList;