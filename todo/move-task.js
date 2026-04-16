/**
 * Move task between To Do lists
 * Graph API has no native move — this creates in target list and deletes from source.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Move task handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveTask(args) {
  const { sourceListId, targetListId, taskId } = args;

  if (!sourceListId || !targetListId || !taskId) {
    return {
      content: [{
        type: "text",
        text: "Source list ID, target list ID, and task ID are all required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Get the existing task
    const sourceEndpoint = `me/todo/lists/${sourceListId}/tasks/${taskId}`;
    const existingTask = await callGraphAPI(accessToken, 'GET', sourceEndpoint);

    // Build new task body preserving all fields
    const newTask = {
      title: existingTask.title,
      importance: existingTask.importance,
      status: existingTask.status,
    };

    if (existingTask.body && existingTask.body.content) {
      newTask.body = {
        content: existingTask.body.content,
        contentType: existingTask.body.contentType || "text"
      };
    }

    if (existingTask.dueDateTime) {
      newTask.dueDateTime = existingTask.dueDateTime;
    }

    if (existingTask.reminderDateTime) {
      newTask.reminderDateTime = existingTask.reminderDateTime;
      newTask.isReminderOn = existingTask.isReminderOn;
    }

    if (existingTask.recurrence) {
      newTask.recurrence = existingTask.recurrence;
    }

    if (existingTask.categories && existingTask.categories.length > 0) {
      newTask.categories = existingTask.categories;
    }

    // Create task in target list
    const targetEndpoint = `me/todo/lists/${targetListId}/tasks`;
    const createdTask = await callGraphAPI(accessToken, 'POST', targetEndpoint, newTask);

    // Delete from source list
    await callGraphAPI(accessToken, 'DELETE', sourceEndpoint);

    return {
      content: [{
        type: "text",
        text: `Task '${createdTask.title}' moved successfully.\nNew ID: ${createdTask.id}`
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
        text: `Error moving task: ${error.message}`
      }]
    };
  }
}

module.exports = handleMoveTask;
