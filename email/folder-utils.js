/**
 * Email folder utilities
 */
const { callGraphAPI } = require("../utils/graph-api");

/**
 * Cache of folder information to reduce API calls
 * Format: { userId: { folderName: { id, path } } }
 */
const folderCache = {};

/**
 * Well-known folder names and their endpoints
 */
const WELL_KNOWN_FOLDERS = {
  inbox: "me/mailFolders/inbox/messages",
  drafts: "me/mailFolders/drafts/messages",
  sent: "me/mailFolders/sentItems/messages",
  deleted: "me/mailFolders/deletedItems/messages",
  junk: "me/mailFolders/junkemail/messages",
  archive: "me/mailFolders/archive/messages",
};

/**
 * Resolve a folder name to its endpoint path
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name to resolve
 * @returns {Promise<string>} - Resolved endpoint path
 */
async function resolveFolderPath(accessToken, folderName) {
  // Default to inbox if no folder specified
  if (!folderName) {
    return WELL_KNOWN_FOLDERS["inbox"];
  }

  // Check if it's a well-known folder (case-insensitive)
  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_FOLDERS[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}"`);
    return WELL_KNOWN_FOLDERS[lowerFolderName];
  }

  const folderId = await getFolderIdByName(accessToken, folderName);
  if (folderId) {
    const path = `me/mailFolders/${folderId}/messages`;
    console.error(`Resolved folder "${folderName}" to path: ${path}`);
    return path;
  }

  throw new Error(`Folder "${folderName}" not found. Check that the folder name and path are correct.`);
}

/**
 * Get the ID of a mail folder by its name, searching recursively through subfolders
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to find (supports "Parent/Child" paths)
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName) {
  try {
    console.error(`Looking for folder with name "${folderName}"`);

    // Support path-based lookup: "Vendors/Anthropic" -> find Anthropic under Vendors
    const pathParts = folderName
      .split("/")
      .map((p) => p.trim())
      .filter((p) => p);
    if (pathParts.length > 1) {
      return await getFolderIdByPath(accessToken, pathParts);
    }

    // Search all folders including subfolders
    const allFolders = await getAllFolders(accessToken);
    const lowerFolderName = folderName.toLowerCase();
    const matchingFolder = allFolders.find(
      (folder) => folder.displayName.toLowerCase() === lowerFolderName,
    );

    if (matchingFolder) {
      console.error(
        `Found folder "${folderName}" with ID: ${matchingFolder.id}`,
      );
      return matchingFolder.id;
    }

    console.error(`No folder found matching "${folderName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    return null;
  }
}

/**
 * Get folder ID by navigating a path array (e.g. ["Vendors", "Anthropic"])
 * @param {string} accessToken - Access token
 * @param {string[]} pathParts - Array of folder name segments
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByPath(accessToken, pathParts) {
  try {
    // Find the first segment anywhere in the folder tree (including subfolders)
    // Use greedy matching: try longest possible segment first to handle folder
    // names that contain "/" (e.g. "Visa/9/15 Renewal" where "9/15 Renewal" is one folder)
    const allFolders = await getAllFolders(accessToken);

    // Try greedy match for the first segment(s) against top-level folders
    let currentId = null;
    let nextIndex = 0;

    for (let length = pathParts.length; length >= 1; length--) {
      const candidate = pathParts.slice(0, length).join("/");
      const lowerCandidate = candidate.toLowerCase();
      const match = allFolders.find(
        (folder) => folder.displayName.toLowerCase() === lowerCandidate,
      );
      if (match) {
        currentId = match.id;
        nextIndex = length;
        console.error(`Matched root segment "${candidate}" (consumed ${length} parts)`);
        break;
      }
    }

    if (!currentId) {
      console.error(`No root path segment found for "${pathParts[0]}"`);
      return null;
    }

    // Navigate remaining segments via childFolders, using greedy matching
    let i = nextIndex;
    while (i < pathParts.length) {
      const response = await callGraphAPI(
        accessToken,
        "GET",
        `me/mailFolders/${currentId}/childFolders`,
        null,
        { $top: 100, $select: "id,displayName" },
      );

      const children = response.value || [];
      let matched = false;

      // Try longest remaining segment first (greedy), then shorter
      for (let length = pathParts.length - i; length >= 1; length--) {
        const candidate = pathParts.slice(i, i + length).join("/");
        const lowerCandidate = candidate.toLowerCase();
        const match = children.find(
          (folder) => folder.displayName.toLowerCase() === lowerCandidate,
        );
        if (match) {
          currentId = match.id;
          console.error(`Matched child segment "${candidate}" (consumed ${length} parts)`);
          i += length;
          matched = true;
          break;
        }
      }

      if (!matched) {
        console.error(
          `Path segment "${pathParts[i]}" not found under folder ID ${currentId}`,
        );
        return null;
      }
    }

    return currentId;
  } catch (error) {
    console.error(`Error navigating folder path: ${error.message}`);
    return null;
  }
}

/**
 * Get all mail folders
 * @param {string} accessToken - Access token
 * @returns {Promise<Array>} - Array of folder objects
 */
async function getAllFolders(accessToken) {
  try {
    const response = await callGraphAPI(
      accessToken,
      "GET",
      "me/mailFolders",
      null,
      {
        $top: 100,
        $select:
          "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount",
      },
    );

    if (!response.value) {
      return [];
    }

    const allFolders = [...response.value];
    await fetchChildFoldersRecursive(accessToken, response.value, allFolders);
    return allFolders;
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

/**
 * Recursively fetch child folders and append them to allFolders
 * @param {string} accessToken - Access token
 * @param {Array} folders - Folders to check for children
 * @param {Array} allFolders - Accumulator array
 */
async function fetchChildFoldersRecursive(accessToken, folders, allFolders) {
  const foldersWithChildren = folders.filter((f) => f.childFolderCount > 0);

  const childResults = await Promise.all(
    foldersWithChildren.map(async (folder) => {
      try {
        const childResponse = await callGraphAPI(
          accessToken,
          "GET",
          `me/mailFolders/${folder.id}/childFolders`,
          null,
          {
            $top: 100,
            $select:
              "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount",
          },
        );
        return childResponse.value || [];
      } catch (error) {
        console.error(
          `Error getting child folders for "${folder.displayName}": ${error.message}`,
        );
        return [];
      }
    }),
  );

  const newFolders = childResults.flat();
  if (newFolders.length > 0) {
    allFolders.push(...newFolders);
    await fetchChildFoldersRecursive(accessToken, newFolders, allFolders);
  }
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  getFolderIdByPath,
  getAllFolders,
  fetchChildFoldersRecursive,
};
