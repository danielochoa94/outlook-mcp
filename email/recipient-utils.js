/**
 * Recipient formatting helpers shared by send/draft/schedule
 */

/**
 * Converts a comma-separated address list into Graph recipient objects
 * @param {string} addresses - Comma-separated email addresses
 * @returns {Array<object>} - Graph recipient objects (empty when nothing usable was given)
 */
function formatRecipients(addresses) {
  if (!addresses) {
    return [];
  }

  return addresses
    .split(',')
    .map(email => ({ emailAddress: { address: email.trim() } }))
    .filter(r => r.emailAddress.address);
}

/**
 * Builds the "Recipients: 2 + 1 CC" summary line used in tool responses
 */
function describeRecipients(toRecipients, ccRecipients = [], bccRecipients = []) {
  return `${toRecipients.length}` +
    `${ccRecipients.length > 0 ? ` + ${ccRecipients.length} CC` : ''}` +
    `${bccRecipients.length > 0 ? ` + ${bccRecipients.length} BCC` : ''}`;
}

module.exports = { formatRecipients, describeRecipients };
