const {
  handleReplyEmail,
  handleForwardEmail,
  composeBody,
  formatSentDate,
  quotedFromText,
  buildDivider,
  replaceDivider,
  mergeRecipients
} = require('../../email/reply');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const FUTURE = '2099-01-31T09:00:00-05:00';

function textOf(response) {
  return response.content[0].text;
}

const RECEIVED = '2026-08-20T19:10:14Z';

function draftFixture(overrides = {}) {
  return {
    id: 'draft-1',
    subject: 'RE: Quarterly numbers',
    body: { contentType: 'html', content: '<html><head></head><body><div>original</div></body></html>' },
    toRecipients: [{ emailAddress: { address: 'sender@example.com' } }],
    ccRecipients: [],
    ...overrides
  };
}

describe('composeBody', () => {
  test('escapes plain text and splices it into the quoted document body', () => {
    const quoted = '<html><head></head><body lang="EN-US"><div>original</div></body></html>';

    const result = composeBody('a < b\nsecond line', { contentType: 'html', content: quoted });

    expect(result.contentType).toBe('html');
    expect(result.content).toBe(
      '<html><head></head><body lang="EN-US">' +
      '<div dir="ltr" style="font-family:Calibri,sans-serif; font-size:11pt">a &lt; b<br>second line</div>' +
      '<div id="appendonsend"></div>' +
      '<div>original</div></body></html>'
    );
  });

  test('falls back to prepending when the quote has no body tag', () => {
    const result = composeBody('hi', { contentType: 'html', content: '<div>original</div>' });

    expect(result.content).toBe(
      '<div dir="ltr" style="font-family:Calibri,sans-serif; font-size:11pt">hi</div>' +
      '<div id="appendonsend"></div><div>original</div>'
    );
  });

  test('stays plain text when there is no divider to rebuild from', () => {
    const result = composeBody('thanks', { contentType: 'text', content: '> original' });

    expect(result).toEqual({ contentType: 'text', content: 'thanks\n\n> original' });
  });

  test('honours isHtml=false against an HTML-looking body', () => {
    const result = composeBody('<html>hi</html>', { contentType: 'text', content: 'original' }, false);

    expect(result).toEqual({ contentType: 'text', content: '<html>hi</html>\n\noriginal' });
  });

  test('wraps a plain quote when the new text is HTML', () => {
    const result = composeBody('<b>hi</b>', { contentType: 'text', content: 'a < b' }, true);

    expect(result).toEqual({
      contentType: 'html',
      content: '<div dir="ltr" style="font-family:Calibri,sans-serif; font-size:11pt"><b>hi</b></div>' +
               '<div id="appendonsend"></div><pre>a &lt; b</pre>'
    });
  });
});

describe('formatSentDate', () => {
  test('formats in local time the way desktop Outlook stamps a divider', () => {
    process.env.TZ = 'America/New_York';

    expect(formatSentDate(RECEIVED)).toBe('Thursday, August 20, 2026 3:10 PM');
  });

  test('returns null for an unusable timestamp', () => {
    expect(formatSentDate(undefined)).toBeNull();
    expect(formatSentDate('not a date')).toBeNull();
  });
});

const ORIGINAL = {
  from: { emailAddress: { name: 'Daniel Ochoa', address: 'daniel@tryeditide.com' } },
  toRecipients: [{ emailAddress: { name: 'Daniel Ochoa', address: 'Daniel@editide.com' } }],
  subject: 'Reply to this',
  receivedDateTime: RECEIVED
};

const GRAPH_DIVIDER =
  '<hr tabindex="-1" style="display:inline-block; width:98%">' +
  '<div id="divRplyFwdMsg" dir="ltr"><font><b>From:</b> x<br>' +
  '<b>Sent:</b> Thursday, 20 August 2026 19:10:14</font> <div>&nbsp;</div></div>';

describe('buildDivider', () => {
  beforeEach(() => {
    process.env.TZ = 'America/New_York';
  });

  test('stamps local time and lists the original recipients', () => {
    const divider = buildDivider(ORIGINAL);

    expect(divider).toContain('border-top:solid #E1E1E1 1.0pt');
    expect(divider).toContain('<b>Sent:</b> Thursday, August 20, 2026 3:10 PM');
    expect(divider).toContain('Daniel Ochoa &lt;daniel@tryeditide.com&gt;');
    expect(divider).toContain('<b>Subject:</b> Reply to this');
    expect(divider).not.toContain('<b>Cc:</b>');
  });

  test('includes a Cc line only when the original had one', () => {
    const divider = buildDivider({
      ...ORIGINAL,
      ccRecipients: [{ emailAddress: { address: 'boss@example.com' } }]
    });

    expect(divider).toContain('<b>Cc:</b> boss@example.com');
  });

  test('escapes markup in names and subjects', () => {
    const divider = buildDivider({ ...ORIGINAL, subject: '<script>alert(1)</script>' });

    expect(divider).toContain('&lt;script&gt;alert(1)&lt;/script&gt;');
    expect(divider).not.toContain('<script>');
  });

  test('gives up when the original lacks a usable timestamp or sender', () => {
    expect(buildDivider({ ...ORIGINAL, receivedDateTime: undefined, sentDateTime: undefined })).toBeNull();
    expect(buildDivider({ ...ORIGINAL, from: undefined })).toBeNull();
    expect(buildDivider(null)).toBeNull();
  });

  test('falls back to sentDateTime when the message was never received', () => {
    const divider = buildDivider({ ...ORIGINAL, receivedDateTime: undefined, sentDateTime: RECEIVED });

    expect(divider).toContain('<b>Sent:</b> Thursday, August 20, 2026 3:10 PM');
  });
});

describe('replaceDivider', () => {
  beforeEach(() => {
    process.env.TZ = 'America/New_York';
  });

  test('replaces the rule and divider block, keeping the quoted body intact', () => {
    const html = `<html><body><div dir="ltr">new</div>${GRAPH_DIVIDER}` +
                 '<div><div dir="ltr">hello <img src="cid:abc"></div></div></body></html>';

    const result = replaceDivider(html, ORIGINAL);

    expect(result).not.toContain('divRplyFwdMsg');
    expect(result).not.toContain('<hr');
    expect(result).toContain('border-top:solid #E1E1E1 1.0pt');
    // The nested <div>&nbsp;</div> must not leave a stray closing tag behind.
    expect(result).toContain('</div><div style="font-size:11.0pt; font-family:&quot;Calibri&quot;,sans-serif">&nbsp;</div>' +
                             '<div><div dir="ltr">hello <img src="cid:abc"></div></div></body>');
  });

  test('leaves a blank line on each side of the divider', () => {
    const html = `<html><body><div dir="ltr">new</div>${GRAPH_DIVIDER}<div>quoted</div></body></html>`;

    const result = replaceDivider(html, ORIGINAL);
    const spacers = result.match(/&nbsp;<\/div>/g) || [];

    expect(spacers).toHaveLength(2);
    expect(result.indexOf('&nbsp;')).toBeLessThan(result.indexOf('border-top'));
    expect(result.lastIndexOf('&nbsp;')).toBeGreaterThan(result.indexOf('border-top'));
  });

  test('leaves the body alone when the original could not be fetched', () => {
    const html = `<html><body>new${GRAPH_DIVIDER}quoted</body></html>`;

    expect(replaceDivider(html, null)).toBe(html);
  });

  test('leaves the body alone when Graph used no divider', () => {
    const html = '<html><body>new<div>quoted</div></body></html>';

    expect(replaceDivider(html, ORIGINAL)).toBe(html);
  });
});

const TEXT_DRAFT = '\n\n________________________________________\n' +
  'From: Daniel Ochoa\nSent: Thursday, 20 August 2026 19:10:14\nTo: daniel@tryeditide.com\n' +
  'Subject: Reply to this\n\nOpening line.\nSecond line.\n';

describe('quotedFromText', () => {
  test('strips the underscore rule and header block, keeping the quoted message', () => {
    expect(quotedFromText(TEXT_DRAFT)).toBe('Opening line.\nSecond line.\n');
  });

  test('returns null when the draft has no rule to split on', () => {
    expect(quotedFromText('just a body')).toBeNull();
    expect(quotedFromText(undefined)).toBeNull();
  });
});

describe('composeBody on a plain-text draft', () => {
  beforeEach(() => {
    process.env.TZ = 'America/New_York';
  });

  test('rebuilds it as HTML with the styled divider', () => {
    const result = composeBody('My reply.', { contentType: 'text', content: TEXT_DRAFT }, undefined, ORIGINAL);

    expect(result.contentType).toBe('html');
    expect(result.content).toContain('<b>Sent:</b> Thursday, August 20, 2026 3:10 PM');
    expect(result.content).toContain('border-top:solid #E1E1E1 1.0pt');
    expect(result.content).toContain('Opening line.<br>Second line.');
    expect(result.content).not.toContain('___');
  });

  test('escapes markup in the quoted text', () => {
    const draft = TEXT_DRAFT.replace('Opening line.', '<script>alert(1)</script>');

    const result = composeBody('hi', { contentType: 'text', content: draft }, undefined, ORIGINAL);

    expect(result.content).toContain('&lt;script&gt;');
    expect(result.content).not.toContain('<script>');
  });

  test('falls back to plain text when the original could not be fetched', () => {
    const result = composeBody('hi', { contentType: 'text', content: TEXT_DRAFT }, undefined, null);

    expect(result.contentType).toBe('text');
    expect(result.content).toContain('___');
  });
});

describe('mergeRecipients', () => {
  test('appends new addresses and drops case-insensitive duplicates', () => {
    const existing = [{ emailAddress: { address: 'a@example.com' } }];
    const added = [{ emailAddress: { address: 'A@Example.com' } }, { emailAddress: { address: 'b@example.com' } }];

    expect(mergeRecipients(existing, added).map(r => r.emailAddress.address))
      .toEqual(['a@example.com', 'b@example.com']);
  });
});

function mockGraph(draft = draftFixture(), original = ORIGINAL) {
  callGraphAPI.mockImplementation((token, method) => {
    if (method === 'GET') return Promise.resolve(original);
    if (method === 'POST') return Promise.resolve(draft);
    return Promise.resolve({});
  });
}

function patchPayload() {
  return callGraphAPI.mock.calls.find(call => call[1] === 'PATCH')[3];
}

describe('handleReplyEmail', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    process.env.TZ = 'America/New_York';
    ensureAuthenticated.mockResolvedValue('token');
  });

  test('creates a reply draft, patches the body and sends it', async () => {
    mockGraph();

    const result = await handleReplyEmail({ id: 'msg-1', body: 'Sounds good.' });

    expect(callGraphAPI).toHaveBeenNthCalledWith(1, 'token', 'POST', 'me/messages/msg-1/createReply');
    expect(patchPayload()).toEqual({
      body: {
        contentType: 'html',
        content: '<html><head></head><body>' +
                 '<div dir="ltr" style="font-family:Calibri,sans-serif; font-size:11pt">Sounds good.</div>' +
                 '<div id="appendonsend"></div><div>original</div></body></html>'
      }
    });
    expect(callGraphAPI).toHaveBeenLastCalledWith('token', 'POST', 'me/messages/draft-1/send');
    expect(textOf(result)).toMatch(/Reply sent/);
  });

  test('rewrites the divider with the Word-style block in local time', async () => {
    mockGraph(draftFixture({
      body: { contentType: 'html', content: `<html><body>${GRAPH_DIVIDER}<div>quoted</div></body></html>` }
    }), ORIGINAL);

    await handleReplyEmail({ id: 'msg-1', body: 'hi' });

    const content = patchPayload().body.content;
    expect(content).toContain('<b>Sent:</b> Thursday, August 20, 2026 3:10 PM');
    expect(content).toContain('border-top:solid #E1E1E1 1.0pt');
    expect(content).not.toContain('divRplyFwdMsg');
  });

  test('sends anyway when the original cannot be fetched for its divider fields', async () => {
    callGraphAPI.mockImplementation((token, method) => {
      if (method === 'GET') return Promise.reject(new Error('Graph API error (status 404)'));
      if (method === 'POST') return Promise.resolve(draftFixture());
      return Promise.resolve({});
    });

    const result = await handleReplyEmail({ id: 'msg-1', body: 'hi' });

    expect(textOf(result)).toMatch(/Reply sent/);
  });

  test('uses createReplyAll when replyAll is set', async () => {
    mockGraph();

    await handleReplyEmail({ id: 'msg-1', body: 'hi', replyAll: true });

    expect(callGraphAPI).toHaveBeenNthCalledWith(1, 'token', 'POST', 'me/messages/msg-1/createReplyAll');
  });

  test('adds extra recipients on top of the prefilled ones', async () => {
    mockGraph();

    await handleReplyEmail({ id: 'msg-1', body: 'hi', to: 'sender@example.com, extra@example.com', cc: 'boss@example.com' });

    const patch = patchPayload();
    expect(patch.toRecipients.map(r => r.emailAddress.address))
      .toEqual(['sender@example.com', 'extra@example.com']);
    expect(patch.ccRecipients.map(r => r.emailAddress.address)).toEqual(['boss@example.com']);
  });

  test('leaves prefilled recipients untouched when none are added', async () => {
    mockGraph();

    await handleReplyEmail({ id: 'msg-1', body: 'hi' });

    expect(patchPayload()).not.toHaveProperty('toRecipients');
  });

  test('attaches the deferred send property when sendAt is given', async () => {
    mockGraph();

    const result = await handleReplyEmail({ id: 'msg-1', body: 'hi', sendAt: FUTURE });

    expect(patchPayload().singleValueExtendedProperties)
      .toEqual([{ id: 'SystemTime 0x3FEF', value: '2099-01-31T14:00:00Z' }]);
    expect(callGraphAPI).toHaveBeenLastCalledWith('token', 'POST', 'me/messages/draft-1/send');
    expect(textOf(result)).toMatch(/Reply scheduled/);
  });

  test('does not send when saveAsDraft is set', async () => {
    mockGraph();

    const result = await handleReplyEmail({ id: 'msg-1', body: 'hi', saveAsDraft: true });

    expect(callGraphAPI.mock.calls.some(call => call[2].endsWith('/send'))).toBe(false);
    expect(textOf(result)).toMatch(/saved as a draft/);
    expect(textOf(result)).toContain('draft-1');
  });

  test('rejects sendAt combined with saveAsDraft', async () => {
    const result = await handleReplyEmail({ id: 'msg-1', body: 'hi', sendAt: FUTURE, saveAsDraft: true });

    expect(textOf(result)).toMatch(/not both/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('rejects a past sendAt before touching Graph', async () => {
    const result = await handleReplyEmail({ id: 'msg-1', body: 'hi', sendAt: '2000-01-01T00:00:00Z' });

    expect(textOf(result)).toMatch(/in the past/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('requires a message id', async () => {
    const result = await handleReplyEmail({ body: 'hi' });

    expect(textOf(result)).toMatch(/Message ID is required/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('reports a missing original message', async () => {
    callGraphAPI.mockRejectedValue(new Error('Graph API error (status 404)'));

    const result = await handleReplyEmail({ id: 'gone', body: 'hi' });

    expect(textOf(result)).toMatch(/No message found with ID gone/);
  });

  test('asks for authentication when the token is missing', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleReplyEmail({ id: 'msg-1', body: 'hi' });

    expect(textOf(result)).toMatch(/Authentication required/);
  });
});

describe('handleForwardEmail', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    process.env.TZ = 'America/New_York';
    ensureAuthenticated.mockResolvedValue('token');
  });

  test('creates a forward draft with the given recipients and sends it', async () => {
    mockGraph(draftFixture({ subject: 'FW: Quarterly numbers', toRecipients: [] }));

    const result = await handleForwardEmail({ id: 'msg-1', to: 'peer@example.com', body: 'FYI' });

    expect(callGraphAPI).toHaveBeenNthCalledWith(1, 'token', 'POST', 'me/messages/msg-1/createForward');
    expect(patchPayload().toRecipients).toEqual([{ emailAddress: { address: 'peer@example.com' } }]);
    expect(callGraphAPI).toHaveBeenLastCalledWith('token', 'POST', 'me/messages/draft-1/send');
    expect(textOf(result)).toMatch(/Forward sent/);
  });

  test('requires a recipient', async () => {
    const result = await handleForwardEmail({ id: 'msg-1', body: 'FYI' });

    expect(textOf(result)).toMatch(/Recipient \(to\) is required/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });
});
