const {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName
} = require('../../email/folder-utils');
const { callGraphAPI } = require('../../utils/graph-api');

jest.mock('../../utils/graph-api');

describe('resolveFolderPath', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    // Mock console.error to avoid cluttering test output
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('well-known folders', () => {
    test('should return inbox endpoint when no folder name is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, null);
      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return inbox endpoint when undefined folder name is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, undefined);
      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return inbox endpoint when empty string is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, '');
      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return correct endpoint for well-known folders', async () => {
      const result = await resolveFolderPath(mockAccessToken, 'drafts');
      expect(result).toBe(WELL_KNOWN_FOLDERS['drafts']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should handle case-insensitive well-known folder names', async () => {
      const result1 = await resolveFolderPath(mockAccessToken, 'INBOX');
      const result2 = await resolveFolderPath(mockAccessToken, 'Drafts');
      const result3 = await resolveFolderPath(mockAccessToken, 'SENT');

      expect(result1).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(result2).toBe(WELL_KNOWN_FOLDERS['drafts']);
      expect(result3).toBe(WELL_KNOWN_FOLDERS['sent']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });
  });

  describe('custom folders', () => {
    test('should resolve custom folder by ID when found', async () => {
      const customFolderId = 'custom-folder-id-123';
      const customFolderName = 'MyCustomFolder';

      // getAllFolders fetches all top-level folders in one call
      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: customFolderId, displayName: customFolderName, childFolderCount: 0 }
        ]
      });

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/mailFolders',
        null,
        {
          $top: 100,
          $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
        }
      );
    });

    test('should resolve custom folder with case-insensitive match', async () => {
      const customFolderId = 'custom-folder-id-456';
      const customFolderName = 'ProjectAlpha';

      // getAllFolders returns all folders; case-insensitive match finds 'projectalpha'
      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'other-id', displayName: 'OtherFolder', childFolderCount: 0 },
          { id: customFolderId, displayName: 'projectalpha', childFolderCount: 0 }
        ]
      });

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
    });

    test('should throw when custom folder is not found', async () => {
      const nonExistentFolder = 'NonExistentFolder';

      // getAllFolders returns folders that don't match
      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'id1', displayName: 'Folder1', childFolderCount: 0 },
          { id: 'id2', displayName: 'Folder2', childFolderCount: 0 }
        ]
      });

      await expect(resolveFolderPath(mockAccessToken, nonExistentFolder))
        .rejects.toThrow('Folder "NonExistentFolder" not found');
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
    });

    test('should throw when API call fails', async () => {
      const customFolderName = 'CustomFolder';

      callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

      await expect(resolveFolderPath(mockAccessToken, customFolderName))
        .rejects.toThrow('Folder "CustomFolder" not found');
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
    });
  });
});

describe('getFolderIdByName', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should return folder ID when exact match is found', async () => {
    const folderId = 'folder-id-123';
    const folderName = 'TestFolder';

    // getAllFolders fetches all folders in one call
    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: folderId, displayName: folderName, childFolderCount: 0 }]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBe(folderId);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/mailFolders',
      null,
      {
        $top: 100,
        $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
      }
    );
  });

  test('should return folder ID when case-insensitive match is found', async () => {
    const folderId = 'folder-id-456';
    const folderName = 'TestFolder';

    // getAllFolders returns all folders; case-insensitive search finds 'testfolder'
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: folderId, displayName: 'testfolder', childFolderCount: 0 }
      ]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBe(folderId);
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  test('should return null when folder is not found', async () => {
    const folderName = 'NonExistentFolder';

    // getAllFolders returns folders that don't match
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: 'id1', displayName: 'OtherFolder', childFolderCount: 0 }
      ]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBeNull();
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  test('should return null when API call fails', async () => {
    const folderName = 'TestFolder';

    callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBeNull();
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });
});
