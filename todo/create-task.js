/**
 * Create To Do task functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Create To Do task handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateTask(args) {
  const { listId, title, body, dueDateTime, importance } = args;

  if (!listId || !title) {
    return {
      content: [{
        type: "text",
        text: "List ID and title are required. Use 'list-todo-lists' to get available list IDs."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/todo/lists/${listId}/tasks`;
    const taskBody = {
      title
    };

    if (body) {
      taskBody.body = {
        content: body,
        contentType: "text"
      };
    }

    if (dueDateTime) {
      taskBody.dueDateTime = {
        dateTime: dueDateTime,
        timeZone: "UTC"
      };
    }

    if (importance) {
      taskBody.importance = importance;
    }

    const response = await callGraphAPI(accessToken, 'POST', endpoint, taskBody);

    const dueInfo = response.dueDateTime
      ? `\nDue: ${response.dueDateTime.dateTime.split('T')[0]}`
      : '';

    return {
      content: [{
        type: "text",
        text: `Task '${response.title}' created successfully.${dueInfo}\nID: ${response.id}`
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
        text: `Error creating task: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateTask;