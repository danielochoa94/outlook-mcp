/**
 * Upload a file to SharePoint
 */
const fs = require('fs');
const path = require('path');
const https = require('https');
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Upload a file to a SharePoint document library
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUploadFile(args) {
  const { siteId, driveId, localPath, destinationPath, folderId } = args;

  if (!siteId || !driveId || !localPath) {
    return {
      content: [{
        type: "text",
        text: "Site ID, Drive ID, and local file path are required."
      }]
    };
  }

  // Verify local file exists
  if (!fs.existsSync(localPath)) {
    return {
      content: [{
        type: "text",
        text: `Local file not found: ${localPath}`
      }]
    };
  }

  const stats = fs.statSync(localPath);
  if (stats.isDirectory()) {
    return {
      content: [{
        type: "text",
        text: "Cannot upload a directory. Please specify a file path."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const fileName = path.basename(localPath);
    const fileSize = stats.size;

    let endpoint;
    if (folderId) {
      endpoint = `sites/${siteId}/drives/${driveId}/items/${folderId}:/${encodeURIComponent(fileName)}:/content`;
    } else if (destinationPath) {
      // destinationPath is the folder path in SharePoint (e.g., "Reports/2026")
      const cleanPath = destinationPath.replace(/^\/+|\/+$/g, '');
      endpoint = `sites/${siteId}/drives/${driveId}/root:/${cleanPath}/${encodeURIComponent(fileName)}:/content`;
    } else {
      endpoint = `sites/${siteId}/drives/${driveId}/root:/${encodeURIComponent(fileName)}:/content`;
    }

    if (fileSize > config.ONEDRIVE_UPLOAD_THRESHOLD) {
      // Large file — use chunked upload session
      const result = await chunkedUpload(accessToken, siteId, driveId, folderId, destinationPath, fileName, localPath, fileSize);
      return {
        content: [{
          type: "text",
          text: `Uploaded "${fileName}" (${formatSize(fileSize)}) via chunked upload.\n   URL: ${result.webUrl}\n   ID: ${result.id}`
        }]
      };
    }

    // Small file — simple PUT upload
    const result = await simpleUpload(accessToken, endpoint, localPath);

    return {
      content: [{
        type: "text",
        text: `Uploaded "${fileName}" (${formatSize(fileSize)}).\n   URL: ${result.webUrl}\n   ID: ${result.id}`
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
        text: `Error uploading file: ${error.message}`
      }]
    };
  }
}

function simpleUpload(accessToken, endpoint, localPath) {
  return new Promise((resolve, reject) => {
    const fileData = fs.readFileSync(localPath);
    const fullUrl = `${config.GRAPH_API_ENDPOINT}${endpoint}`;

    const options = {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/octet-stream',
        'Content-Length': fileData.length
      }
    };

    const request = https.request(fullUrl, options, (response) => {
      let data = '';
      response.on('data', (chunk) => { data += chunk; });
      response.on('end', () => {
        if (response.statusCode >= 200 && response.statusCode < 300) {
          try {
            resolve(JSON.parse(data));
          } catch (error) {
            reject(new Error(`Error parsing upload response: ${error.message}`));
          }
        } else if (response.statusCode === 401) {
          reject(new Error('UNAUTHORIZED'));
        } else {
          reject(new Error(`Upload failed with status ${response.statusCode}: ${data}`));
        }
      });
    });

    request.on('error', (error) => {
      reject(new Error(`Network error during upload: ${error.message}`));
    });

    request.write(fileData);
    request.end();
  });
}

async function chunkedUpload(accessToken, siteId, driveId, folderId, destinationPath, fileName, localPath, fileSize) {
  // Create upload session
  let sessionEndpoint;
  if (folderId) {
    sessionEndpoint = `sites/${siteId}/drives/${driveId}/items/${folderId}:/${encodeURIComponent(fileName)}:/createUploadSession`;
  } else if (destinationPath) {
    const cleanPath = destinationPath.replace(/^\/+|\/+$/g, '');
    sessionEndpoint = `sites/${siteId}/drives/${driveId}/root:/${cleanPath}/${encodeURIComponent(fileName)}:/createUploadSession`;
  } else {
    sessionEndpoint = `sites/${siteId}/drives/${driveId}/root:/${encodeURIComponent(fileName)}:/createUploadSession`;
  }

  const session = await callGraphAPI(accessToken, 'POST', sessionEndpoint, {
    item: {
      "@microsoft.graph.conflictBehavior": "replace"
    }
  });

  const uploadUrl = session.uploadUrl;
  const chunkSize = 3.25 * 1024 * 1024; // 3.25 MB chunks (must be multiple of 320 KiB)
  const fileBuffer = fs.readFileSync(localPath);
  let offset = 0;

  while (offset < fileSize) {
    const end = Math.min(offset + chunkSize, fileSize);
    const chunk = fileBuffer.slice(offset, end);

    const result = await uploadChunk(uploadUrl, chunk, offset, end - 1, fileSize);

    if (result.id) {
      // Upload complete
      return result;
    }

    offset = end;
  }

  throw new Error('Chunked upload completed without receiving final response');
}

function uploadChunk(uploadUrl, chunk, rangeStart, rangeEnd, totalSize) {
  return new Promise((resolve, reject) => {
    const options = {
      method: 'PUT',
      headers: {
        'Content-Length': chunk.length,
        'Content-Range': `bytes ${rangeStart}-${rangeEnd}/${totalSize}`
      }
    };

    const request = https.request(uploadUrl, options, (response) => {
      let data = '';
      response.on('data', (part) => { data += part; });
      response.on('end', () => {
        if (response.statusCode >= 200 && response.statusCode < 300) {
          try {
            resolve(JSON.parse(data));
          } catch {
            resolve({});
          }
        } else {
          reject(new Error(`Chunk upload failed with status ${response.statusCode}: ${data}`));
        }
      });
    });

    request.on('error', (error) => {
      reject(new Error(`Network error during chunk upload: ${error.message}`));
    });

    request.write(chunk);
    request.end();
  });
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

module.exports = handleUploadFile;
