/**
 * Move all contents of one folder to another
 */
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getFolderIdByName } = require('../email/folder-utils');

/**
 * Move folder contents handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveFolderContents(args) {
  const sourceFolder = args.sourceFolder || '';
  const targetFolder = args.targetFolder || '';

  if (!sourceFolder) {
    return {
      content: [{
        type: "text",
        text: "Source folder is required."
      }]
    };
  }

  if (!targetFolder) {
    return {
      content: [{
        type: "text",
        text: "Target folder is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Resolve both folder IDs
    const sourceFolderId = await getFolderIdByName(accessToken, sourceFolder);
    if (!sourceFolderId) {
      return {
        content: [{
          type: "text",
          text: `Source folder "${sourceFolder}" not found.`
        }]
      };
    }

    const targetFolderId = await getFolderIdByName(accessToken, targetFolder);
    if (!targetFolderId) {
      return {
        content: [{
          type: "text",
          text: `Target folder "${targetFolder}" not found.`
        }]
      };
    }

    // Get all message IDs from the source folder (only fetch id to minimize payload)
    const response = await callGraphAPIPaginated(
      accessToken,
      'GET',
      `me/mailFolders/${sourceFolderId}/messages`,
      { $select: 'id', $top: 50 },
      0 // 0 = get all
    );

    const messages = response.value || [];

    if (messages.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No emails found in "${sourceFolder}". Folder is already empty.`
        }]
      };
    }

    // Move messages in batches
    const batchSize = 20;
    let successful = 0;
    let failed = 0;
    const errors = [];

    for (let i = 0; i < messages.length; i += batchSize) {
      const batch = messages.slice(i, i + batchSize);

      const results = await Promise.allSettled(
        batch.map(message =>
          callGraphAPI(
            accessToken,
            'POST',
            `me/messages/${message.id}/move`,
            { destinationId: targetFolderId }
          )
        )
      );

      for (const result of results) {
        if (result.status === 'fulfilled') {
          successful++;
        } else {
          failed++;
          if (errors.length < 3) {
            errors.push(result.reason.message);
          }
        }
      }

      console.error(`Move progress: ${successful + failed}/${messages.length} processed`);
    }

    // Build result message
    let message = `Moved ${successful} email(s) from "${sourceFolder}" to "${targetFolder}".`;

    if (failed > 0) {
      message += `\n${failed} email(s) failed to move.`;
      for (const error of errors) {
        message += `\n- ${error}`;
      }
      if (failed > errors.length) {
        message += `\n...and ${failed - errors.length} more errors.`;
      }
    }

    return {
      content: [{
        type: "text",
        text: message
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
        text: `Error moving folder contents: ${error.message}`
      }]
    };
  }
}

module.exports = handleMoveFolderContents;
