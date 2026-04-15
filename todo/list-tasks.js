/**
 * List To Do tasks functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List To Do tasks handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListTasks(args) {
  const { listId, status, count } = args;

  if (!listId) {
    return {
      content: [{
        type: "text",
        text: "List ID is required. Use 'list-todo-lists' to get available list IDs."
      }]
    };
  }

  const maxCount = Math.min(count || 25, config.MAX_RESULT_COUNT);

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/todo/lists/${listId}/tasks`;
    const queryParams = {
      $top: maxCount,
      $orderby: 'createdDateTime desc'
    };

    if (status === 'completed') {
      queryParams.$filter = 'status eq \'completed\'';
    } else if (status === 'active') {
      queryParams.$filter = 'status ne \'completed\'';
    }

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No tasks found in this list."
        }]
      };
    }

    const taskOutput = response.value.map((task, index) => {
      const statusIcon = task.status === 'completed' ? '[x]' : '[ ]';
      const importance = task.importance === 'high' ? ' !' : '';
      const dueDate = task.dueDateTime
        ? `\n   Due: ${task.dueDateTime.dateTime.split('T')[0]}`
        : '';
      const body = task.body?.content ? `\n   Note: ${task.body.content.substring(0, 100)}` : '';

      return `${index + 1}. ${statusIcon}${importance} ${task.title}${dueDate}${body}\n   ID: ${task.id}`;
    }).join("\n");

    const statusLabel = status ? ` (${status})` : '';
    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} tasks${statusLabel}:\n\n${taskOutput}`
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
        text: `Error listing tasks: ${error.message}`
      }]
    };
  }
}

module.exports = handleListTasks;