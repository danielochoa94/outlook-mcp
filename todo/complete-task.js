/**
 * Complete To Do task functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Complete To Do task handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCompleteTask(args) {
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
    const updateBody = {
      status: "completed"
    };

    const response = await callGraphAPI(accessToken, 'PATCH', endpoint, updateBody);

    return {
      content: [{
        type: "text",
        text: `Task '${response.title}' marked as completed.`
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
        text: `Error completing task: ${error.message}`
      }]
    };
  }
}

module.exports = handleCompleteTask;