/**
 * Send email functionality
 *
 * One tool covers the three ways a new message can leave: straight out through sendMail, parked in
 * Drafts, or held by Exchange until a deferred send time. Reply and forward take the same options,
 * so "when it goes out" is never a separate tool.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { formatRecipients, describeRecipients } = require('./recipient-utils');
const { parseSendTime, DEFERRED_SEND_TIME_PROPERTY } = require('./schedule');

function textResponse(text) {
  return { content: [{ type: "text", text }] };
}

/**
 * Send email handler: sends now, saves a draft, or schedules a deferred send
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSendEmail(args) {
  const {
    to, cc, bcc, subject, body, importance = 'normal',
    saveToSentItems = true, isHtml, sendAt, saveAsDraft = false
  } = args || {};

  if (!to) {
    return textResponse("Recipient (to) is required.");
  }

  if (!subject) {
    return textResponse("Subject is required.");
  }

  if (!body) {
    return textResponse("Body content is required.");
  }

  if (sendAt && saveAsDraft) {
    return textResponse("Use either sendAt or saveAsDraft, not both: a saved draft is never sent, a scheduled one always is.");
  }

  let sendTime = null;
  if (sendAt) {
    sendTime = parseSendTime(sendAt);
    if (sendTime.error) {
      return textResponse(sendTime.error);
    }
  }

  try {
    const accessToken = await ensureAuthenticated();

    const toRecipients = formatRecipients(to);
    const ccRecipients = formatRecipients(cc);
    const bccRecipients = formatRecipients(bcc);

    if (toRecipients.length === 0) {
      return textResponse("No valid recipient addresses were found in 'to'. A recipient is required.");
    }

    const contentType = isHtml === true ? 'html' :
                        isHtml === false ? 'text' :
                        body.toLowerCase().includes('<html') ? 'html' : 'text';

    const message = {
      subject,
      body: { contentType, content: body },
      toRecipients,
      ccRecipients: ccRecipients.length > 0 ? ccRecipients : undefined,
      bccRecipients: bccRecipients.length > 0 ? bccRecipients : undefined,
      importance
    };

    const summary = `Subject: ${subject}\nRecipients: ${describeRecipients(toRecipients, ccRecipients, bccRecipients)}`;

    if (!sendTime && !saveAsDraft) {
      await callGraphAPI(accessToken, 'POST', 'me/sendMail', { message, saveToSentItems });

      return textResponse(`Email sent successfully!\n\n${summary}\nMessage Length: ${body.length} characters`);
    }

    if (sendTime) {
      message.singleValueExtendedProperties = [{ id: DEFERRED_SEND_TIME_PROPERTY, value: sendTime.utc }];
    }

    const draft = await callGraphAPI(accessToken, 'POST', 'me/messages', message);

    if (!draft || !draft.id) {
      return textResponse(`Draft creation returned no message id, so the email was not ${sendTime ? 'scheduled' : 'saved'}.`);
    }

    if (saveAsDraft) {
      return textResponse(
        `Draft created successfully!\n\n${summary}\nDraft ID: ${draft.id}\n\n` +
        `It sits in Drafts until you send it from Outlook.`
      );
    }

    // Sending a deferred draft queues it; Exchange releases it at the deferred time.
    await callGraphAPI(accessToken, 'POST', `me/messages/${draft.id}/send`);

    return textResponse(
      `Email scheduled successfully!\n\n${summary}\n` +
      `Sends at: ${sendTime.date.toISOString()} (UTC)\nMessage ID: ${draft.id}\n\n` +
      `It stays in Drafts until then. Use 'cancel-scheduled-email' with that ID to stop it.`
    );
  } catch (error) {
    if (error.message === 'Authentication required') {
      return textResponse("Authentication required. Please use the 'authenticate' tool first.");
    }

    if (error.message && error.message.includes('status 403')) {
      return textResponse(
        "Microsoft Graph denied the request (403). The token likely lacks Mail.ReadWrite scope. " +
        "Re-authenticate with force=true to refresh consent, then try again."
      );
    }

    return textResponse(`Error sending email: ${error.message}`);
  }
}

module.exports = handleSendEmail;
