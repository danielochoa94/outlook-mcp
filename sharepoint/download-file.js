/**
 * Download a file from SharePoint
 */
const fs = require('fs');
const path = require('path');
const https = require('https');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Download a file from a SharePoint document library
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDownloadFile(args) {
  const { siteId, driveId, itemId, savePath } = args;

  if (!siteId || !driveId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "Site ID, Drive ID, and Item ID are required. Use 'list-sharepoint-files' to find them."
      }]
    };
  }

  if (!savePath) {
    return {
      content: [{
        type: "text",
        text: "Save path is required. Provide a local file path to save the downloaded file."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Get file metadata first to know the name and size
    const metadata = await callGraphAPI(
      accessToken,
      'GET',
      `sites/${siteId}/drives/${driveId}/items/${itemId}`,
      null,
      { $select: 'name,size,file' }
    );

    if (!metadata.file) {
      return {
        content: [{
          type: "text",
          text: "The specified item is not a file. Only files can be downloaded."
        }]
      };
    }

    // Get the download URL
    const downloadUrl = await getDownloadUrl(accessToken, siteId, driveId, itemId);

    // Resolve save path - if it's a directory, append the filename
    let resolvedPath = savePath;
    try {
      const stats = fs.statSync(savePath);
      if (stats.isDirectory()) {
        resolvedPath = path.join(savePath, metadata.name);
      }
    } catch {
      // Path doesn't exist yet, use as-is (will create the file)
    }

    // Ensure parent directory exists
    const parentDirectory = path.dirname(resolvedPath);
    if (!fs.existsSync(parentDirectory)) {
      fs.mkdirSync(parentDirectory, { recursive: true });
    }

    // Download the file
    await downloadToFile(downloadUrl, resolvedPath);

    return {
      content: [{
        type: "text",
        text: `Downloaded "${metadata.name}" (${formatSize(metadata.size)}) to:\n${resolvedPath}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error downloading file: ${error.message}`
      }]
    };
  }
}

function getDownloadUrl(accessToken, siteId, driveId, itemId) {
  return new Promise((resolve, reject) => {
    const endpoint = `sites/${siteId}/drives/${driveId}/items/${itemId}/content`;
    const fullUrl = `https://graph.microsoft.com/v1.0/${endpoint}`;

    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    };

    const request = https.request(fullUrl, options, (response) => {
      if (response.statusCode === 302 && response.headers.location) {
        resolve(response.headers.location);
      } else if (response.statusCode === 200) {
        // Sometimes the API returns the content directly with a download URL in metadata
        let data = '';
        response.on('data', (chunk) => { data += chunk; });
        response.on('end', () => {
          try {
            const json = JSON.parse(data);
            if (json['@microsoft.graph.downloadUrl']) {
              resolve(json['@microsoft.graph.downloadUrl']);
            } else {
              reject(new Error('No download URL found in response'));
            }
          } catch {
            reject(new Error('Unexpected response format'));
          }
        });
      } else if (response.statusCode === 401) {
        reject(new Error('UNAUTHORIZED'));
      } else {
        let data = '';
        response.on('data', (chunk) => { data += chunk; });
        response.on('end', () => {
          reject(new Error(`Download request failed with status ${response.statusCode}: ${data}`));
        });
      }
    });

    request.on('error', (error) => {
      reject(new Error(`Network error: ${error.message}`));
    });

    request.end();
  });
}

function downloadToFile(downloadUrl, filePath) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(filePath);

    https.get(downloadUrl, (response) => {
      if (response.statusCode === 200) {
        response.pipe(file);
        file.on('finish', () => {
          file.close(resolve);
        });
      } else if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        // Follow redirect
        file.close();
        fs.unlinkSync(filePath);
        downloadToFile(response.headers.location, filePath).then(resolve).catch(reject);
      } else {
        file.close();
        fs.unlinkSync(filePath);
        reject(new Error(`Download failed with status ${response.statusCode}`));
      }
    }).on('error', (error) => {
      file.close();
      fs.unlinkSync(filePath);
      reject(new Error(`Download error: ${error.message}`));
    });
  });
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

module.exports = handleDownloadFile;
