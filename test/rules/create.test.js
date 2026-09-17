const handleCreateRule = require('../../rules/create');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleCreateRule recipientContains', () => {
  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'new-rule-id' });
  });

  test('accepts recipientContains as the only condition', async () => {
    // Arrange
    const args = {
      name: 'Block siteboost CC',
      recipientContains: 'siteboost@wildppcagency.com',
      markAsRead: true,
      sequence: 1
    };

    // Act
    const result = await handleCreateRule(args);

    // Assert
    expect(result.content[0].text).toContain('Successfully created rule "Block siteboost CC"');
  });

  test('sends recipientContains to Graph as a string array', async () => {
    // Arrange
    const args = {
      name: 'Block siteboost CC',
      recipientContains: 'siteboost@wildppcagency.com',
      markAsRead: true,
      sequence: 1
    };

    // Act
    await handleCreateRule(args);

    // Assert
    const rule = callGraphAPI.mock.calls[0][3];
    expect(rule.conditions.recipientContains).toEqual(['siteboost@wildppcagency.com']);
  });

  test('splits and trims a comma-separated recipientContains list', async () => {
    // Arrange
    const args = {
      name: 'Block siteboost CC',
      recipientContains: 'siteboost@wildppcagency.com , wildppcagency.com',
      markAsRead: true,
      sequence: 1
    };

    // Act
    await handleCreateRule(args);

    // Assert
    const rule = callGraphAPI.mock.calls[0][3];
    expect(rule.conditions.recipientContains).toEqual([
      'siteboost@wildppcagency.com',
      'wildppcagency.com'
    ]);
  });

  test('omits recipientContains when only blank entries are supplied', async () => {
    // Arrange
    const args = {
      name: 'Blank recipients',
      recipientContains: ' , ',
      markAsRead: true,
      sequence: 1
    };

    // Act
    const result = await handleCreateRule(args);

    // Assert
    expect(result.content[0].text).toContain('At least one condition is required');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('names recipientContains in the missing-condition error', async () => {
    // Arrange
    const args = { name: 'No conditions', markAsRead: true, sequence: 1 };

    // Act
    const result = await handleCreateRule(args);

    // Assert
    expect(result.content[0].text).toContain('recipientContains');
  });

  test('combines recipientContains with fromAddresses', async () => {
    // Arrange
    const args = {
      name: 'Both conditions',
      fromAddresses: 'contato@imigrante.email-bot.chat',
      recipientContains: 'siteboost@wildppcagency.com',
      markAsRead: true,
      sequence: 1
    };

    // Act
    await handleCreateRule(args);

    // Assert
    const rule = callGraphAPI.mock.calls[0][3];
    expect(rule.conditions.recipientContains).toEqual(['siteboost@wildppcagency.com']);
    expect(rule.conditions.fromAddresses).toEqual([
      { emailAddress: { address: 'contato@imigrante.email-bot.chat' } }
    ]);
  });
});
