/**
 * Update To Do task functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Update To Do task handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateTask(args) {
  const { listId, taskId, title, body, status, dueDateTime, importance } = args;

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
    const updateBody = {};

    if (title !== undefined) {
      updateBody.title = title;
    }

    if (body !== undefined) {
      updateBody.body = {
        content: body,
        contentType: "text"
      };
    }

    if (status !== undefined) {
      updateBody.status = status;
    }

    if (dueDateTime !== undefined) {
      updateBody.dueDateTime = dueDateTime
        ? { dateTime: dueDateTime, timeZone: "UTC" }
        : null;
    }

    if (importance !== undefined) {
      updateBody.importance = importance;
    }

    if (Object.keys(updateBody).length === 0) {
      return {
        content: [{
          type: "text",
          text: "No fields provided to update. Specify at least one of: title, body, status, dueDateTime, importance."
        }]
      };
    }

    const response = await callGraphAPI(accessToken, 'PATCH', endpoint, updateBody);

    return {
      content: [{
        type: "text",
        text: `Task '${response.title}' updated successfully.`
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
        text: `Error updating task: ${error.message}`
      }]
    };
  }
}

module.exports = handleUpdateTask;