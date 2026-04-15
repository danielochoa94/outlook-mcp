/**
 * Delete To Do task functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Delete To Do task handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteTask(args) {
  const { listId, taskId } = args;

  if (!listId || !taskId) {
    return {
      content: [{
        type: "text",
        text: "List ID and task ID are required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/todo/lists/${listId}/tasks/${taskId}`;

    await callGraphAPI(accessToken, 'DELETE', endpoint);

    return {
      content: [{
        type: "text",
        text: `Task with ID ${taskId} has been successfully deleted.`
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
        text: `Error deleting task: ${error.message}`
      }]
    };
  }
}

module.exports = handleDeleteTask;