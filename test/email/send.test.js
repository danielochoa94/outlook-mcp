const handleSendEmail = require('../../email/send');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { DEFERRED_SEND_TIME_PROPERTY } = require('../../email/schedule');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const FUTURE = '2099-01-31T09:00:00-05:00';

function textOf(response) {
  return response.content[0].text;
}

const BASE = { to: 'a@example.com', subject: 'Hello', body: 'Plain body' };

describe('handleSendEmail', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('token');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('sends immediately through sendMail', async () => {
    callGraphAPI.mockResolvedValueOnce({});

    const result = await handleSendEmail({ ...BASE, cc: 'c@example.com' });

    expect(callGraphAPI).toHaveBeenCalledTimes(1);
    const [, method, path, payload] = callGraphAPI.mock.calls[0];
    expect(method).toBe('POST');
    expect(path).toBe('me/sendMail');
    expect(payload.message.body).toEqual({ contentType: 'text', content: 'Plain body' });
    expect(payload.message.toRecipients).toHaveLength(1);
    expect(textOf(result)).toMatch(/Email sent successfully/);
  });

  test('auto-detects html bodies', async () => {
    callGraphAPI.mockResolvedValueOnce({});

    await handleSendEmail({ ...BASE, body: '<HTML><body>hi</body></HTML>' });

    expect(callGraphAPI.mock.calls[0][3].message.body.contentType).toBe('html');
  });

  test('sendAt creates a deferred draft and hands it to the transport', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'draft-1' })
      .mockResolvedValueOnce({});

    const result = await handleSendEmail({ ...BASE, sendAt: FUTURE });

    expect(callGraphAPI).toHaveBeenCalledTimes(2);
    const [, , createPath, message] = callGraphAPI.mock.calls[0];
    expect(createPath).toBe('me/messages');
    expect(message.singleValueExtendedProperties).toEqual([
      { id: DEFERRED_SEND_TIME_PROPERTY, value: '2099-01-31T14:00:00Z' }
    ]);
    expect(callGraphAPI.mock.calls[1][2]).toBe('me/messages/draft-1/send');
    expect(textOf(result)).toMatch(/Email scheduled/);
    expect(textOf(result)).toMatch(/cancel-scheduled-email/);
  });

  test('saveAsDraft creates the draft without sending it', async () => {
    callGraphAPI.mockResolvedValueOnce({ id: 'draft-2', subject: 'Hello' });

    const result = await handleSendEmail({ ...BASE, saveAsDraft: true });

    expect(callGraphAPI).toHaveBeenCalledTimes(1);
    expect(callGraphAPI.mock.calls[0][2]).toBe('me/messages');
    expect(callGraphAPI.mock.calls[0][3].singleValueExtendedProperties).toBeUndefined();
    expect(textOf(result)).toMatch(/Draft created/);
    expect(textOf(result)).toMatch(/draft-2/);
  });

  test('rejects sendAt combined with saveAsDraft', async () => {
    const result = await handleSendEmail({ ...BASE, sendAt: FUTURE, saveAsDraft: true });

    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(textOf(result)).toMatch(/not both/);
  });

  test('rejects a sendAt in the past', async () => {
    const result = await handleSendEmail({ ...BASE, sendAt: '2000-01-01T00:00:00Z' });

    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(textOf(result)).toMatch(/in the past/);
  });

  test('rejects a sendAt with no timezone', async () => {
    const result = await handleSendEmail({ ...BASE, sendAt: '2099-01-31T09:00:00' });

    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(textOf(result)).toMatch(/no timezone/);
  });

  test.each(['to', 'subject', 'body'])('requires %s', async (field) => {
    const args = { ...BASE };
    delete args[field];

    const result = await handleSendEmail(args);

    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(textOf(result)).toMatch(new RegExp(field, 'i'));
  });

  test('reports when no recipient address parses', async () => {
    const result = await handleSendEmail({ ...BASE, to: '   ' });

    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(textOf(result)).toMatch(/required/i);
  });

  test('explains a 403 on draft creation as a missing scope', async () => {
    callGraphAPI.mockRejectedValueOnce(new Error('Graph API error (status 403)'));

    const result = await handleSendEmail({ ...BASE, saveAsDraft: true });

    expect(textOf(result)).toMatch(/Mail\.ReadWrite/);
  });

  test('asks for authentication when the token is missing', async () => {
    ensureAuthenticated.mockRejectedValueOnce(new Error('Authentication required'));

    const result = await handleSendEmail(BASE);

    expect(textOf(result)).toMatch(/authenticate/i);
  });
});
