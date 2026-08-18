/**
 * Scheduled (deferred) email send functionality
 *
 * Exchange holds a message in Drafts/Outbox until PidTagDeferredSendTime passes, so scheduling is
 * "create a draft carrying that property, then send it" - no client needs to stay running.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { formatRecipients, describeRecipients } = require('./recipient-utils');

const DEFERRED_SEND_TIME_PROPERTY = 'SystemTime 0x3FEF';
const SCHEDULED_FOLDERS = ['drafts', 'outbox'];
const TIMEZONE_SUFFIX = /(Z|[+-]\d{2}:?\d{2})$/i;

/**
 * Validates an ISO 8601 datetime and normalises it to the UTC form Graph expects
 * @param {string} value - Requested send time, must carry an explicit timezone
 * @returns {{ utc: string, date: Date } | { error: string }}
 */
function parseSendTime(value) {
  if (!value) {
    return { error: "sendAt is required, e.g. '2025-01-31T09:00:00-05:00' or '2025-01-31T14:00:00Z'." };
  }

  if (!TIMEZONE_SUFFIX.test(value.trim())) {
    return { error: `sendAt (${value}) has no timezone. Add 'Z' for UTC or an offset such as '-05:00' so the send time is unambiguous.` };
  }

  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return { error: `sendAt (${value}) is not a valid ISO 8601 datetime.` };
  }

  if (date.getTime() <= Date.now()) {
    return { error: `sendAt (${value}) is in the past. Pick a future time.` };
  }

  return { utc: date.toISOString().replace(/\.\d{3}Z$/, 'Z'), date };
}

function authErrorResponse() {
  return {
    content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }]
  };
}

function textResponse(text) {
  return { content: [{ type: "text", text }] };
}

/**
 * Schedule email handler: creates a deferred draft and hands it to the transport
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleScheduleEmail(args) {
  const { to, cc, bcc, subject, body, sendAt, importance = 'normal', isHtml } = args || {};

  if (!to) {
    return textResponse("Recipient (to) is required.");
  }

  if (!subject) {
    return textResponse("Subject is required.");
  }

  if (!body) {
    return textResponse("Body content is required.");
  }

  const sendTime = parseSendTime(sendAt);
  if (sendTime.error) {
    return textResponse(sendTime.error);
  }

  try {
    const accessToken = await ensureAuthenticated();

    const toRecipients = formatRecipients(to);
    const ccRecipients = formatRecipients(cc);
    const bccRecipients = formatRecipients(bcc);

    if (toRecipients.length === 0) {
      return textResponse("No valid recipient addresses were found in 'to'.");
    }

    const contentType = isHtml === true ? 'html' :
                        isHtml === false ? 'text' :
                        body.toLowerCase().includes('<html') ? 'html' : 'text';

    const messageObject = {
      subject,
      body: { contentType, content: body },
      toRecipients,
      ccRecipients: ccRecipients.length > 0 ? ccRecipients : undefined,
      bccRecipients: bccRecipients.length > 0 ? bccRecipients : undefined,
      importance,
      singleValueExtendedProperties: [
        { id: DEFERRED_SEND_TIME_PROPERTY, value: sendTime.utc }
      ]
    };

    const draft = await callGraphAPI(accessToken, 'POST', 'me/messages', messageObject);

    if (!draft || !draft.id) {
      return textResponse("Draft creation returned no message id, so the email was not scheduled.");
    }

    // Sending a deferred draft queues it; Exchange releases it at the deferred time.
    await callGraphAPI(accessToken, 'POST', `me/messages/${draft.id}/send`);

    return textResponse(
      `Email scheduled successfully!\n\n` +
      `Subject: ${subject}\n` +
      `Recipients: ${describeRecipients(toRecipients, ccRecipients, bccRecipients)}\n` +
      `Sends at: ${sendTime.date.toISOString()} (UTC)\n` +
      `Message ID: ${draft.id}\n\n` +
      `It stays in Drafts until then. Use 'cancel-scheduled-email' with that ID to stop it.`
    );
  } catch (error) {
    if (error.message === 'Authentication required') {
      return authErrorResponse();
    }

    return textResponse(`Error scheduling email: ${error.message}`);
  }
}

/**
 * Lists messages still waiting on a deferred send time
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListScheduledEmails(args) {
  const { count = 25 } = args || {};

  try {
    const accessToken = await ensureAuthenticated();

    const queryParams = {
      $top: Math.min(Math.max(parseInt(count, 10) || 25, 1), 50),
      $select: 'id,subject,toRecipients,createdDateTime',
      $expand: `singleValueExtendedProperties($filter=id eq '${DEFERRED_SEND_TIME_PROPERTY}')`
    };

    const scheduled = [];
    for (const folder of SCHEDULED_FOLDERS) {
      const response = await callGraphAPI(
        accessToken, 'GET', `me/mailFolders/${folder}/messages`, null, queryParams
      );

      for (const message of response.value || []) {
        // Graph echoes the property id lower-cased ("SystemTime 0x3fef"), so match case-insensitively.
        const property = (message.singleValueExtendedProperties || [])
          .find(p => p.id && p.id.toLowerCase() === DEFERRED_SEND_TIME_PROPERTY.toLowerCase());

        if (property) {
          scheduled.push({ folder, message, sendAt: property.value });
        }
      }
    }

    if (scheduled.length === 0) {
      return textResponse("No scheduled emails found.");
    }

    scheduled.sort((a, b) => new Date(a.sendAt) - new Date(b.sendAt));

    const lines = scheduled.map(({ folder, message, sendAt }, index) => {
      const recipients = (message.toRecipients || [])
        .map(r => r.emailAddress.address)
        .join(', ') || '(none)';

      return `${index + 1}. ${message.subject || '(no subject)'}\n` +
             `   To: ${recipients}\n` +
             `   Sends at: ${sendAt}\n` +
             `   Folder: ${folder}\n` +
             `   ID: ${message.id}`;
    });

    return textResponse(`Found ${scheduled.length} scheduled email(s):\n\n${lines.join('\n\n')}`);
  } catch (error) {
    if (error.message === 'Authentication required') {
      return authErrorResponse();
    }

    return textResponse(`Error listing scheduled emails: ${error.message}`);
  }
}

/**
 * Cancels a pending scheduled send by deleting the queued message
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCancelScheduledEmail(args) {
  const { id } = args || {};

  if (!id) {
    return textResponse("Message ID is required. Use 'list-scheduled-emails' to find it.");
  }

  try {
    const accessToken = await ensureAuthenticated();

    await callGraphAPI(accessToken, 'DELETE', `me/messages/${id}`);

    return textResponse(
      `Scheduled email cancelled.\n\nMessage ID: ${id}\nThe queued message was moved to Deleted Items, so it will not be sent.`
    );
  } catch (error) {
    if (error.message === 'Authentication required') {
      return authErrorResponse();
    }

    if (error.message && error.message.includes('status 404')) {
      return textResponse(`No message found with ID ${id}. It may have already been sent or cancelled.`);
    }

    return textResponse(`Error cancelling scheduled email: ${error.message}`);
  }
}

module.exports = {
  handleScheduleEmail,
  handleListScheduledEmails,
  handleCancelScheduledEmail,
  parseSendTime,
  DEFERRED_SEND_TIME_PROPERTY
};
