/**
 * To Do module for M365 Assistant MCP server
 */
const handleListTodoLists = require('./list-lists');
const handleCreateTodoList = require('./create-list');
const handleDeleteTodoList = require('./delete-list');
const handleListTasks = require('./list-tasks');
const handleCreateTask = require('./create-task');
const handleUpdateTask = require('./update-task');
const handleCompleteTask = require('./complete-task');
const handleDeleteTask = require('./delete-task');

// To Do tool definitions
const todoTools = [
  {
    name: "list-todo-lists",
    description: "Lists all Microsoft To Do task lists",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleListTodoLists
  },
  {
    name: "create-todo-list",
    description: "Creates a new Microsoft To Do task list",
    inputSchema: {
      type: "object",
      properties: {
        displayName: {
          type: "string",
          description: "Name of the task list to create"
        }
      },
      required: ["displayName"]
    },
    handler: handleCreateTodoList
  },
  {
    name: "delete-todo-list",
    description: "Deletes a Microsoft To Do task list",
    inputSchema: {
      type: "object",
      properties: {
        listId: {
          type: "string",
          description: "The ID of the task list to delete"
        }
      },
      required: ["listId"]
    },
    handler: handleDeleteTodoList
  },
  {
    name: "list-tasks",
    description: "Lists tasks in a Microsoft To Do task list. Use status filter to show all, active, or completed tasks.",
    inputSchema: {
      type: "object",
      properties: {
        listId: {
          type: "string",
          description: "The ID of the task list"
        },
        status: {
          type: "string",
          enum: ["all", "active", "completed"],
          description: "Filter by status (default: all)"
        },
        count: {
          type: "number",
          description: "Number of tasks to retrieve (default: 25, max: 50)"
        }
      },
      required: ["listId"]
    },
    handler: handleListTasks
  },
  {
    name: "create-task",
    description: "Creates a new task in a Microsoft To Do task list",
    inputSchema: {
      type: "object",
      properties: {
        listId: {
          type: "string",
          description: "The ID of the task list to add the task to"
        },
        title: {
          type: "string",
          description: "Title of the task"
        },
        body: {
          type: "string",
          description: "Optional note/body content for the task"
        },
        dueDateTime: {
          type: "string",
          description: "Optional due date in ISO 8601 format (e.g. 2026-04-20T00:00:00)"
        },
        importance: {
          type: "string",
          enum: ["low", "normal", "high"],
          description: "Task importance level (default: normal)"
        }
      },
      required: ["listId", "title"]
    },
    handler: handleCreateTask
  },
  {
    name: "update-task",
    description: "Updates an existing task in a Microsoft To Do task list. Provide only the fields you want to change.",
    inputSchema: {
      type: "object",
      properties: {
        listId: {
          type: "string",
          description: "The ID of the task list"
        },
        taskId: {
          type: "string",
          description: "The ID of the task to update"
        },
        title: {
          type: "string",
          description: "New title for the task"
        },
        body: {
          type: "string",
          description: "New note/body content for the task"
        },
        status: {
          type: "string",
          enum: ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"],
          description: "New status for the task"
        },
        dueDateTime: {
          type: "string",
          description: "New due date in ISO 8601 format, or empty string to clear"
        },
        importance: {
          type: "string",
          enum: ["low", "normal", "high"],
          description: "New importance level"
        }
      },
      required: ["listId", "taskId"]
    },
    handler: handleUpdateTask
  },
  {
    name: "complete-task",
    description: "Marks a Microsoft To Do task as completed",
    inputSchema: {
      type: "object",
      properties: {
        listId: {
          type: "string",
          description: "The ID of the task list"
        },
        taskId: {
          type: "string",
          description: "The ID of the task to complete"
        }
      },
      required: ["listId", "taskId"]
    },
    handler: handleCompleteTask
  },
  {
    name: "delete-task",
    description: "Deletes a task from a Microsoft To Do task list",
    inputSchema: {
      type: "object",
      properties: {
        listId: {
          type: "string",
          description: "The ID of the task list"
        },
        taskId: {
          type: "string",
          description: "The ID of the task to delete"
        }
      },
      required: ["listId", "taskId"]
    },
    handler: handleDeleteTask
  }
];

module.exports = {
  todoTools,
  handleListTodoLists,
  handleCreateTodoList,
  handleDeleteTodoList,
  handleListTasks,
  handleCreateTask,
  handleUpdateTask,
  handleCompleteTask,
  handleDeleteTask
};