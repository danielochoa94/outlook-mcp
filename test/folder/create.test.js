const handleCreateFolder = require('../../folder/create');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { getFolderIdByName } = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils');

describe('handleCreateFolder', () => {
  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    getFolderIdByName.mockClear();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
  });

  test('creates a root-level folder when no parent is specified', async () => {
    // No duplicate found, creation succeeds
    callGraphAPI.mockResolvedValueOnce({ value: [] });
    callGraphAPI.mockResolvedValueOnce({ id: 'new-folder-id' });

    const result = await handleCreateFolder({ name: 'MyFolder' });

    expect(getFolderIdByName).not.toHaveBeenCalled();
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      1,
      'dummy_access_token',
      'GET',
      'me/mailFolders',
      null,
      { $filter: "displayName eq 'MyFolder'", $select: 'id,displayName' },
    );
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      2,
      'dummy_access_token',
      'POST',
      'me/mailFolders',
      { displayName: 'MyFolder' },
    );
    expect(result.content[0].text).toContain('Successfully created folder "MyFolder" at the root level');
  });

  test('creates a subfolder under a parent', async () => {
    getFolderIdByName.mockResolvedValue('parent-id');
    // No duplicate found under parent, creation succeeds
    callGraphAPI.mockResolvedValueOnce({ value: [] });
    callGraphAPI.mockResolvedValueOnce({ id: 'new-folder-id' });

    const result = await handleCreateFolder({ name: 'Legal', parentFolder: 'Reference' });

    expect(getFolderIdByName).toHaveBeenCalledWith('dummy_access_token', 'Reference');
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      1,
      'dummy_access_token',
      'GET',
      'me/mailFolders/parent-id/childFolders',
      null,
      { $filter: "displayName eq 'Legal'", $select: 'id,displayName' },
    );
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      2,
      'dummy_access_token',
      'POST',
      'me/mailFolders/parent-id/childFolders',
      { displayName: 'Legal' },
    );
    expect(result.content[0].text).toContain('Successfully created folder "Legal" inside "Reference"');
  });

  test('does NOT block creation when same name exists in a different parent', async () => {
    // "Legal" exists under Inbox but we are creating under Reference — the scoped
    // check returns empty, so creation should proceed.
    getFolderIdByName.mockResolvedValue('reference-id');
    callGraphAPI.mockResolvedValueOnce({ value: [] }); // no duplicate under Reference
    callGraphAPI.mockResolvedValueOnce({ id: 'new-folder-id' });

    const result = await handleCreateFolder({ name: 'Legal', parentFolder: 'Reference' });

    expect(result.content[0].text).toContain('Successfully created folder "Legal" inside "Reference"');
  });

  test('blocks creation when same name already exists in the same parent', async () => {
    getFolderIdByName.mockResolvedValue('reference-id');
    callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'existing-id', displayName: 'Legal' }] });

    const result = await handleCreateFolder({ name: 'Legal', parentFolder: 'Reference' });

    expect(callGraphAPI).toHaveBeenCalledTimes(1); // only the check, not the creation
    expect(result.content[0].text).toContain('already exists inside "Reference"');
  });

  test('blocks creation when same name already exists at root level', async () => {
    callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'existing-id', displayName: 'Archive' }] });

    const result = await handleCreateFolder({ name: 'Archive' });

    expect(callGraphAPI).toHaveBeenCalledTimes(1);
    expect(result.content[0].text).toContain('already exists at the root level');
  });

  test('returns error when parent folder is not found', async () => {
    getFolderIdByName.mockResolvedValue(null);

    const result = await handleCreateFolder({ name: 'Legal', parentFolder: 'NonExistent' });

    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(result.content[0].text).toContain('Parent folder "NonExistent" not found');
  });

  test('returns error when folder name is missing', async () => {
    const result = await handleCreateFolder({});

    expect(ensureAuthenticated).not.toHaveBeenCalled();
    expect(result.content[0].text).toBe('Folder name is required.');
  });

  test('handles authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleCreateFolder({ name: 'MyFolder' });

    expect(result.content[0].text).toContain('Authentication required');
  });

  test('escapes single quotes in folder name for OData filter', async () => {
    callGraphAPI.mockResolvedValueOnce({ value: [] });
    callGraphAPI.mockResolvedValueOnce({ id: 'new-folder-id' });

    await handleCreateFolder({ name: "Daniel's Folder" });

    expect(callGraphAPI).toHaveBeenNthCalledWith(
      1,
      'dummy_access_token',
      'GET',
      'me/mailFolders',
      null,
      { $filter: "displayName eq 'Daniel''s Folder'", $select: 'id,displayName' },
    );
  });
});
