const { handleListRules } = require('../../rules/list');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleListRules recipientContains', () => {
  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
  });

  test('displays recipientContains in detailed output', async () => {
    // Arrange
    callGraphAPI.mockResolvedValue({
      value: [{
        displayName: 'Block siteboost CC',
        isEnabled: true,
        sequence: 1,
        conditions: { recipientContains: ['siteboost@wildppcagency.com'] },
        actions: { markAsRead: true }
      }]
    });

    // Act
    const result = await handleListRules({ includeDetails: true });

    // Assert
    expect(result.content[0].text).toContain('To or CC contains: "siteboost@wildppcagency.com"');
  });
});
