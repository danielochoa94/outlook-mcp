/**
 * Reply and forward functionality
 *
 * Everything goes through Graph's createReply/createReplyAll/createForward drafts: Exchange fills in
 * the In-Reply-To/References headers and the quoted original, so the result actually threads in
 * Outlook. We then patch our text on top of the quote and either send it, defer it, or leave it in
 * Drafts.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { formatRecipients, describeRecipients } = require('./recipient-utils');
const { parseSendTime, DEFERRED_SEND_TIME_PROPERTY } = require('./schedule');

function textResponse(text) {
  return { content: [{ type: "text", text }] };
}

function escapeHtml(value) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function isHtmlBody(body, isHtml) {
  if (isHtml === true) return true;
  if (isHtml === false) return false;
  return typeof body === 'string' && body.toLowerCase().includes('<html');
}

// Outlook's own compose font; without it the new text renders in the client's default serif.
const COMPOSE_FONT = 'font-family:Calibri,sans-serif; font-size:11pt';

// Graph's own divider is a <hr> + #divRplyFwdMsg block stamped in UTC; desktop Outlook emits a
// Word-style bordered block in local time, which is what people expect to see in a thread.
const QUOTE_FONT = 'font-size:11.0pt; font-family:&quot;Calibri&quot;,sans-serif';
const DIVIDER_ID = '<div id="divRplyFwdMsg"';
// Outlook leaves a blank line on each side of the divider; Graph's <hr> and trailing &nbsp; block
// supplied that spacing, so replacing them means supplying it ourselves.
const SPACER = `<div style="${QUOTE_FONT}">&nbsp;</div>`;
// Graph's plain-text drafts separate the quote with a rule of underscores followed by a header block.
const TEXT_RULE = /\n_{5,}[ \t]*\r?\n/;

/**
 * Formats a timestamp the way desktop Outlook stamps a reply divider, in the host's timezone
 * @param {string} isoDateTime - Original message's arrival time
 * @returns {string|null} - e.g. "Thursday, August 20, 2026 3:10 PM", or null if unparseable
 */
function formatSentDate(isoDateTime) {
  const date = new Date(isoDateTime);

  if (!isoDateTime || isNaN(date.getTime())) {
    return null;
  }

  const parts = new Intl.DateTimeFormat('en-US', {
    weekday: 'long', month: 'long', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit'
  }).formatToParts(date).reduce((acc, part) => ({ ...acc, [part.type]: part.value }), {});

  return `${parts.weekday}, ${parts.month} ${parts.day}, ${parts.year} ${parts.hour}:${parts.minute} ${parts.dayPeriod}`;
}

function describeAddress(recipient) {
  const { name, address } = (recipient && recipient.emailAddress) || {};

  if (!address) {
    return '';
  }

  return name ? `${escapeHtml(name)} &lt;${escapeHtml(address)}&gt;` : escapeHtml(address);
}

/**
 * Builds the bordered From/Sent/To/Subject block desktop Outlook puts above a quote
 * @param {object} original - Original message (from, toRecipients, ccRecipients, subject, receivedDateTime)
 * @returns {string|null} - Divider markup, or null when the message lacks the fields to build one
 */
function buildDivider(original) {
  const sent = formatSentDate(original && (original.receivedDateTime || original.sentDateTime));

  if (!sent || !original.from) {
    return null;
  }

  const to = (original.toRecipients || []).map(describeAddress).filter(Boolean).join('; ');
  const cc = (original.ccRecipients || []).map(describeAddress).filter(Boolean).join('; ');

  const lines = [
    `<b>Sent:</b> ${sent}`,
    `<b>To:</b> ${to}`,
    cc ? `<b>Cc:</b> ${cc}` : null,
    `<b>Subject:</b> ${escapeHtml(original.subject || '')}`
  ].filter(Boolean).join('<br>');

  return '<div style="border:none; border-top:solid #E1E1E1 1.0pt; padding:3.0pt 0in 0in 0in">' +
         `<p class="MsoNormal"><b><span style="${QUOTE_FONT}">From:</span></b>` +
         `<span style="${QUOTE_FONT}"> ${describeAddress(original.from)} <br>${lines}</span></p></div>`;
}

/**
 * Finds the end of the element opened at `start`, counting nested <div> tags
 * @param {string} html - Markup to scan
 * @param {number} start - Index of the opening <div
 * @returns {number} - Index just past the matching </div>, or -1 when unbalanced
 */
function endOfDiv(html, start) {
  const tags = /<\/?div\b/gi;
  tags.lastIndex = start;

  let depth = 0;
  let tag;

  while ((tag = tags.exec(html))) {
    depth += tag[0][1] === '/' ? -1 : 1;

    if (depth === 0) {
      const close = html.indexOf('>', tag.index);
      return close === -1 ? -1 : close + 1;
    }
  }

  return -1;
}

/**
 * Swaps Graph's divider block for the Word-style one, leaving the quoted body (and its inline
 * image references) exactly as Graph wired them
 * @param {string} html - Composed reply body
 * @param {object|null} original - Original message, or null when it could not be fetched
 * @returns {string} - The body, with the divider replaced when possible
 */
function replaceDivider(html, original) {
  const divider = buildDivider(original);
  const start = html.indexOf(DIVIDER_ID);

  if (!divider || start === -1) {
    return html;
  }

  const end = endOfDiv(html, start);
  if (end === -1) {
    return html;
  }

  // Graph puts a horizontal rule immediately before the block; the bordered divider replaces it.
  const rule = html.lastIndexOf('<hr', start);
  const from = rule === -1 ? start : rule;

  return html.slice(0, from) + SPACER + divider + SPACER + html.slice(end);
}

/**
 * Splices content in just after the opening <body> tag, keeping the document well-formed
 * @param {string} html - Full HTML document from the draft
 * @param {string} injected - Markup to place at the top of the body
 * @returns {string} - The document with injected content in place
 */
function insertIntoBody(html, injected) {
  const bodyTag = /<body\b[^>]*>/i.exec(html);

  if (!bodyTag) {
    return injected + html;
  }

  const at = bodyTag.index + bodyTag[0].length;
  return html.slice(0, at) + injected + html.slice(at);
}

/**
 * Pulls the quoted message out of a plain-text draft, dropping Graph's underscore rule and its
 * header block so the styled divider can take their place
 * @param {string} text - Draft body Graph produced
 * @returns {string|null} - The quoted message alone, or null when the draft has no rule
 */
function quotedFromText(text) {
  const rule = TEXT_RULE.exec(text || '');

  if (!rule) {
    return null;
  }

  const afterRule = text.slice(rule.index + rule[0].length);
  const headerEnd = afterRule.search(/\r?\n\s*\r?\n/);

  return headerEnd === -1 ? '' : afterRule.slice(headerEnd).replace(/^\s+/, '');
}

function asHtmlBlock(text) {
  return `<div style="${QUOTE_FONT}">${escapeHtml(text).replace(/\n/g, '<br>')}</div>`;
}

/**
 * Puts the new text above the quoted original Graph put in the draft
 * @param {string} body - Text the user is adding
 * @param {object} draftBody - The draft's existing body ({ contentType, content })
 * @param {boolean|undefined} isHtml - Explicit content type for the new text
 * @param {object|null} original - Original message, used to rebuild a plain-text draft as HTML
 * @returns {object} - Graph itemBody
 */
function composeBody(body, draftBody, isHtml, original) {
  const quotedType = ((draftBody && draftBody.contentType) || 'text').toLowerCase();
  const quoted = (draftBody && draftBody.content) || '';
  const newIsHtml = isHtmlBody(body, isHtml);
  const inner = newIsHtml ? body : escapeHtml(body).replace(/\n/g, '<br>');
  const injected = `<div dir="ltr" style="${COMPOSE_FONT}">${inner}</div><div id="appendonsend"></div>`;

  if (quotedType !== 'html' && !newIsHtml) {
    // Graph replies to a plain-text message in kind; rebuild it as HTML so the reply looks the same
    // whatever format the original used.
    const divider = buildDivider(original);
    const quotedText = quotedFromText(quoted);

    if (!divider || quotedText === null) {
      return { contentType: 'text', content: quoted ? `${body}\n\n${quoted}` : body };
    }

    return {
      contentType: 'html',
      content: injected + SPACER + divider + SPACER + asHtmlBlock(quotedText)
    };
  }

  if (!quoted) {
    return { contentType: 'html', content: injected };
  }

  const content = quotedType === 'html'
    ? insertIntoBody(quoted, injected)
    : `${injected}<pre>${escapeHtml(quoted)}</pre>`;

  return { contentType: 'html', content };
}

/**
 * Adds recipients to the ones Graph prefilled, ignoring duplicates
 */
function mergeRecipients(existing = [], added = []) {
  const seen = new Set(existing.map(r => r.emailAddress.address.toLowerCase()));
  const merged = [...existing];

  for (const recipient of added) {
    const address = recipient.emailAddress.address.toLowerCase();
    if (!seen.has(address)) {
      seen.add(address);
      merged.push(recipient);
    }
  }

  return merged;
}

/**
 * Shared implementation for reply, reply-all and forward
 * @param {string} action - Graph draft-creating action name
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function createAndDeliver(action, args) {
  const { id, body = '', to, cc, bcc, importance, isHtml, sendAt, saveAsDraft = false } = args || {};
  const verb = action === 'createForward' ? 'forward' : 'reply';

  if (!id) {
    return textResponse(`Message ID is required. Use 'list-emails' or 'search-emails' to find the message to ${verb} to.`);
  }

  if (sendAt && saveAsDraft) {
    return textResponse("Use either sendAt or saveAsDraft, not both: a saved draft is never sent, a scheduled one always is.");
  }

  const addedTo = formatRecipients(to);
  if (action === 'createForward' && addedTo.length === 0) {
    return textResponse("Recipient (to) is required when forwarding.");
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

    // The draft never carries the original's send time, so fetch it in parallel to restamp the divider.
    const [draft, original] = await Promise.all([
      callGraphAPI(accessToken, 'POST', `me/messages/${id}/${action}`),
      callGraphAPI(accessToken, 'GET', `me/messages/${id}`, null, {
        $select: 'from,toRecipients,ccRecipients,subject,receivedDateTime,sentDateTime'
      }).catch(() => null)
    ]);

    if (!draft || !draft.id) {
      return textResponse(`Graph returned no draft for the ${verb}, so nothing was sent.`);
    }

    const composed = composeBody(body, draft.body, isHtml, original);
    if (composed.contentType === 'html') {
      composed.content = replaceDivider(composed.content, original);
    }

    const update = { body: composed };

    const toRecipients = mergeRecipients(draft.toRecipients, addedTo);
    const ccRecipients = mergeRecipients(draft.ccRecipients, formatRecipients(cc));
    const bccRecipients = mergeRecipients(draft.bccRecipients, formatRecipients(bcc));

    if (toRecipients.length > (draft.toRecipients || []).length) update.toRecipients = toRecipients;
    if (ccRecipients.length > (draft.ccRecipients || []).length) update.ccRecipients = ccRecipients;
    if (bccRecipients.length > (draft.bccRecipients || []).length) update.bccRecipients = bccRecipients;
    if (importance) update.importance = importance;

    if (sendTime) {
      update.singleValueExtendedProperties = [{ id: DEFERRED_SEND_TIME_PROPERTY, value: sendTime.utc }];
    }

    await callGraphAPI(accessToken, 'PATCH', `me/messages/${draft.id}`, update);

    const summary = `Subject: ${draft.subject || '(no subject)'}\n` +
                    `Recipients: ${describeRecipients(toRecipients, ccRecipients, bccRecipients)}`;

    if (saveAsDraft) {
      return textResponse(
        `${verb === 'forward' ? 'Forward' : 'Reply'} saved as a draft.\n\n${summary}\nDraft ID: ${draft.id}\n\n` +
        `It sits in Drafts, threaded to the original, until you send it from Outlook.`
      );
    }

    await callGraphAPI(accessToken, 'POST', `me/messages/${draft.id}/send`);

    if (sendTime) {
      return textResponse(
        `${verb === 'forward' ? 'Forward' : 'Reply'} scheduled.\n\n${summary}\n` +
        `Sends at: ${sendTime.date.toISOString()} (UTC)\nMessage ID: ${draft.id}\n\n` +
        `It stays in Drafts until then. Use 'cancel-scheduled-email' with that ID to stop it.`
      );
    }

    return textResponse(`${verb === 'forward' ? 'Forward' : 'Reply'} sent.\n\n${summary}`);
  } catch (error) {
    if (error.message === 'Authentication required') {
      return textResponse("Authentication required. Please use the 'authenticate' tool first.");
    }

    if (error.message && error.message.includes('status 404')) {
      return textResponse(`No message found with ID ${id}. It may have been moved or deleted.`);
    }

    return textResponse(`Error creating ${verb}: ${error.message}`);
  }
}

async function handleReplyEmail(args) {
  const action = (args || {}).replyAll ? 'createReplyAll' : 'createReply';
  return createAndDeliver(action, args);
}

async function handleForwardEmail(args) {
  return createAndDeliver('createForward', args);
}

module.exports = {
  handleReplyEmail,
  handleForwardEmail,
  composeBody,
  quotedFromText,
  insertIntoBody,
  formatSentDate,
  buildDivider,
  replaceDivider,
  mergeRecipients
};
