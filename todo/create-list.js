/**
 * Create To Do task list functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Create To Do task list handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateTodoList(args) {
  const { displayName } = args;

  if (!displayName) {
    return {
      content: [{
        type: "text",
        text: "Display name is required to create a task list."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = 'me/todo/lists';
    const body = { displayName };

    const response = await callGraphAPI(accessToken, 'POST', endpoint, body);

    return {
      content: [{
        type: "text",
        text: `Task list '${response.displayName}' created successfully.\nID: ${response.id}`
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
        text: `Error creating task list: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateTodoList;