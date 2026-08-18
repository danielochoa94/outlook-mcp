const {
  handleScheduleEmail,
  handleListScheduledEmails,
  handleCancelScheduledEmail,
  parseSendTime,
  DEFERRED_SEND_TIME_PROPERTY
} = require('../../email/schedule');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const FUTURE = '2099-01-31T09:00:00-05:00';

function textOf(response) {
  return response.content[0].text;
}

describe('parseSendTime', () => {
  test('normalises an offset time to UTC without milliseconds', () => {
    expect(parseSendTime(FUTURE).utc).toBe('2099-01-31T14:00:00Z');
  });

  test('rejects a time without a timezone', () => {
    expect(parseSendTime('2099-01-31T09:00:00').error).toMatch(/no timezone/);
  });

  test('rejects an unparseable time', () => {
    expect(parseSendTime('2099-13-45T00:00:00Z').error).toMatch(/not a valid ISO 8601/);
  });

  test('rejects a past time', () => {
    expect(parseSendTime('2000-01-01T00:00:00Z').error).toMatch(/in the past/);
  });

  test('rejects a missing time', () => {
    expect(parseSendTime(undefined).error).toMatch(/required/);
  });
});

describe('handleScheduleEmail', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('token');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('creates a deferred draft and sends it', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'draft-1' })
      .mockResolvedValueOnce({});

    const result = await handleScheduleEmail({
      to: 'a@example.com, b@example.com',
      cc: 'c@example.com',
      subject: 'Standup notes',
      body: 'See attached.',
      sendAt: FUTURE
    });

    const [, method, path, payload] = callGraphAPI.mock.calls[0];
    expect(method).toBe('POST');
    expect(path).toBe('me/messages');
    expect(payload.singleValueExtendedProperties).toEqual([
      { id: DEFERRED_SEND_TIME_PROPERTY, value: '2099-01-31T14:00:00Z' }
    ]);
    expect(payload.toRecipients).toHaveLength(2);
    expect(payload.body.contentType).toBe('text');

    expect(callGraphAPI.mock.calls[1][2]).toBe('me/messages/draft-1/send');
    expect(textOf(result)).toMatch(/scheduled successfully/);
    expect(textOf(result)).toMatch(/draft-1/);
  });

  test('detects HTML bodies and honours an explicit isHtml override', async () => {
    callGraphAPI.mockResolvedValue({ id: 'draft-1' });

    await handleScheduleEmail({
      to: 'a@example.com', subject: 's', body: '<html><p>hi</p></html>', sendAt: FUTURE
    });
    expect(callGraphAPI.mock.calls[0][3].body.contentType).toBe('html');

    callGraphAPI.mockClear();
    await handleScheduleEmail({
      to: 'a@example.com', subject: 's', body: '<html><p>hi</p></html>', sendAt: FUTURE, isHtml: false
    });
    expect(callGraphAPI.mock.calls[0][3].body.contentType).toBe('text');
  });

  test('does not call the API when the send time is invalid', async () => {
    const result = await handleScheduleEmail({
      to: 'a@example.com', subject: 's', body: 'b', sendAt: '2099-01-31T09:00:00'
    });

    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(textOf(result)).toMatch(/no timezone/);
  });

  test('requires to, subject and body', async () => {
    expect(textOf(await handleScheduleEmail({ subject: 's', body: 'b', sendAt: FUTURE }))).toMatch(/Recipient/);
    expect(textOf(await handleScheduleEmail({ to: 'a@x.com', body: 'b', sendAt: FUTURE }))).toMatch(/Subject/);
    expect(textOf(await handleScheduleEmail({ to: 'a@x.com', subject: 's', sendAt: FUTURE }))).toMatch(/Body/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('reports send failures', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'draft-1' })
      .mockRejectedValueOnce(new Error('API call failed with status 403: nope'));

    const result = await handleScheduleEmail({
      to: 'a@example.com', subject: 's', body: 'b', sendAt: FUTURE
    });

    expect(textOf(result)).toMatch(/Error scheduling email: .*403/);
  });
});

describe('handleListScheduledEmails', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('token');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('lists only messages carrying a deferred send time, earliest first', async () => {
    callGraphAPI
      .mockResolvedValueOnce({
        value: [
          {
            id: 'later',
            subject: 'Later',
            toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
            singleValueExtendedProperties: [{ id: DEFERRED_SEND_TIME_PROPERTY, value: '2099-02-01T10:00:00Z' }]
          },
          { id: 'plain-draft', subject: 'Not scheduled' }
        ]
      })
      .mockResolvedValueOnce({
        value: [
          {
            id: 'sooner',
            subject: 'Sooner',
            toRecipients: [{ emailAddress: { address: 'b@example.com' } }],
            singleValueExtendedProperties: [{ id: DEFERRED_SEND_TIME_PROPERTY, value: '2099-01-01T10:00:00Z' }]
          }
        ]
      });

    const text = textOf(await handleListScheduledEmails({}));

    expect(text).toMatch(/Found 2 scheduled/);
    expect(text.indexOf('Sooner')).toBeLessThan(text.indexOf('Later'));
    expect(text).not.toMatch(/Not scheduled/);

    const queryParams = callGraphAPI.mock.calls[0][4];
    expect(queryParams.$expand).toContain(DEFERRED_SEND_TIME_PROPERTY);
  });

  test('matches the property id case-insensitively as Graph returns it', async () => {
    callGraphAPI
      .mockResolvedValueOnce({
        value: [{
          id: 'lower',
          subject: 'Lowercased id',
          toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
          singleValueExtendedProperties: [{ id: 'SystemTime 0x3fef', value: '2099-01-01T10:00:00Z' }]
        }]
      })
      .mockResolvedValueOnce({ value: [] });

    expect(textOf(await handleListScheduledEmails({}))).toMatch(/Lowercased id/);
  });

  test('reports when nothing is scheduled', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });

    expect(textOf(await handleListScheduledEmails({}))).toBe('No scheduled emails found.');
  });

  test('clamps count to the API page limit', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });

    await handleListScheduledEmails({ count: 500 });

    expect(callGraphAPI.mock.calls[0][4].$top).toBe(50);
  });
});

describe('handleCancelScheduledEmail', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('token');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('deletes the queued message', async () => {
    callGraphAPI.mockResolvedValue({});

    const result = await handleCancelScheduledEmail({ id: 'draft-1' });

    expect(callGraphAPI).toHaveBeenCalledWith('token', 'DELETE', 'me/messages/draft-1');
    expect(textOf(result)).toMatch(/cancelled/);
  });

  test('requires an id', async () => {
    expect(textOf(await handleCancelScheduledEmail({}))).toMatch(/Message ID is required/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('explains a 404 as already sent or cancelled', async () => {
    callGraphAPI.mockRejectedValue(new Error('API call failed with status 404: not found'));

    expect(textOf(await handleCancelScheduledEmail({ id: 'gone' }))).toMatch(/already been sent or cancelled/);
  });
});
