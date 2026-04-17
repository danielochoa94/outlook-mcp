/**
 * Improved search emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');

/**
 * Search emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchEmails(args) {
  const folder = args.folder;
  const requestedCount = args.count || 10;
  const skip = args.skip || 0;
  const query = args.query || '';
  const from = args.from || '';
  const to = args.to || '';
  const subject = args.subject || '';
  const hasAttachments = args.hasAttachments;
  const unreadOnly = args.unreadOnly;

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Search across all folders when no folder is specified; otherwise scope to the requested folder
    const endpoint = folder
      ? await resolveFolderPath(accessToken, folder)
      : "me/messages";
    console.error(`Using endpoint: ${endpoint} for folder: ${folder || '(all folders)'}`);
    
    // Execute progressive search with pagination
    const response = await progressiveSearch(
      endpoint,
      accessToken,
      { query, from, to, subject },
      { hasAttachments, unreadOnly },
      requestedCount,
      skip
    );
    
    return formatSearchResults(response, skip);
  } catch (error) {
    // Handle authentication errors
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    // General error response
    return {
      content: [{ 
        type: "text", 
        text: `Error searching emails: ${error.message}`
      }]
    };
  }
}

/**
 * Execute a search with progressively simpler fallback strategies
 * @param {string} endpoint - API endpoint
 * @param {string} accessToken - Access token
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} maxCount - Maximum number of results to retrieve
 * @param {number} skip - Number of results to skip for pagination
 * @returns {Promise<object>} - Search results
 */
async function progressiveSearch(endpoint, accessToken, searchTerms, filterTerms, maxCount, skip = 0) {
  // Track search strategies attempted
  const searchAttempts = [];
  
  // 1. Try combined search (most specific)
  try {
    const params = buildSearchParams(searchTerms, filterTerms, Math.min(50, maxCount), skip);
    console.error("Attempting combined search with params:", params);
    searchAttempts.push("combined-search");
    
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, params, maxCount);
    if (response.value && response.value.length > 0) {
      console.error(`Combined search successful: found ${response.value.length} results`);
      return response;
    }
  } catch (error) {
    console.error(`Combined search failed: ${error.message}`);
  }
  
  // 2. Try each search term individually, starting with most specific
  const searchPriority = ['subject', 'from', 'to', 'query'];
  
  for (const term of searchPriority) {
    if (searchTerms[term]) {
      try {
        console.error(`Attempting search with only ${term}: "${searchTerms[term]}"`);
        searchAttempts.push(`single-term-${term}`);
        
        // For single term search, only use $search with that term
        // Graph API does not support $orderby or $filter with $search
        const simplifiedParams = {
          $top: Math.min(50, maxCount),
          $select: config.EMAIL_SELECT_FIELDS
        };

        if (skip > 0) {
          simplifiedParams.$skip = skip;
        }
        
        // Build KQL terms for search
        const kqlParts = [];
        
        // Add the search term in the appropriate KQL syntax
        if (term === 'query') {
          // General query doesn't need a prefix
          kqlParts.push(searchTerms[term]);
        } else {
          // Specific field searches use field:value syntax
          kqlParts.push(`${term}:${searchTerms[term]}`);
        }
        
        // Add boolean filters as KQL (can't use $filter with $search)
        addBooleanFiltersAsKQL(kqlParts, filterTerms);
        
        simplifiedParams.$search = `"${kqlParts.join(' ')}"`;
        
        const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, simplifiedParams, maxCount);
        if (response.value && response.value.length > 0) {
          console.error(`Search with ${term} successful: found ${response.value.length} results`);
          return response;
        }
      } catch (error) {
        console.error(`Search with ${term} failed: ${error.message}`);
      }
    }
  }
  
  // 3. Boolean-filters-only path: run only when the caller supplied no search terms
  //    and at least one boolean filter. Otherwise return empty — the user's search
  //    terms genuinely matched nothing, and falling back to recent emails would be misleading.
  const hasSearchTerm = Boolean(
    searchTerms.query || searchTerms.from || searchTerms.to || searchTerms.subject
  );
  const hasBooleanFilter = filterTerms.hasAttachments === true || filterTerms.unreadOnly === true;

  if (!hasSearchTerm && hasBooleanFilter) {
    console.error("Attempting search with only boolean filters");
    searchAttempts.push("boolean-filters-only");

    const filterOnlyParams = {
      $top: Math.min(50, maxCount),
      $select: config.EMAIL_SELECT_FIELDS,
      $orderby: 'receivedDateTime desc'
    };

    if (skip > 0) {
      filterOnlyParams.$skip = skip;
    }

    addBooleanFilters(filterOnlyParams, filterTerms);

    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, filterOnlyParams, maxCount);
    console.error(`Boolean filter search found ${response.value?.length || 0} results`);
    return response;
  }

  // No fallback to recent emails: if search terms matched nothing, return empty.
  console.error(`All search strategies returned empty; returning empty result (attempts: ${searchAttempts.join(', ')})`);
  return { value: [] };
}

/**
 * Build search parameters from search terms and filter terms
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} count - Maximum number of results
 * @param {number} skip - Number of results to skip
 * @returns {object} - Query parameters
 */
function buildSearchParams(searchTerms, filterTerms, count, skip = 0) {
  const params = {
    $top: count,
    $select: config.EMAIL_SELECT_FIELDS
  };

  if (skip > 0) {
    params.$skip = skip;
  }
  
  // Handle search terms
  const kqlTerms = [];
  
  if (searchTerms.query) {
    // General query doesn't need a prefix
    kqlTerms.push(searchTerms.query);
  }
  
  if (searchTerms.subject) {
    kqlTerms.push(`subject:\"${searchTerms.subject}\"`);
  }

  if (searchTerms.from) {
    kqlTerms.push(`from:\"${searchTerms.from}\"`);
  }

  if (searchTerms.to) {
    kqlTerms.push(`to:\"${searchTerms.to}\"`);
  }
  
  // Add $search if we have any search terms
  if (kqlTerms.length > 0) {
    // Graph API does not support $orderby or $filter with $search
    // Move boolean filters into KQL syntax instead
    addBooleanFiltersAsKQL(kqlTerms, filterTerms);
    params.$search = `"${kqlTerms.join(' ')}"`;
  } else {
    // No search terms — safe to use $orderby and $filter
    params.$orderby = 'receivedDateTime desc';
    addBooleanFilters(params, filterTerms);
  }
  
  return params;
}

/**
 * Add boolean filters to query parameters as OData $filter
 * Only use when $search is NOT present (they conflict in Graph API)
 * @param {object} params - Query parameters
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 */
function addBooleanFilters(params, filterTerms) {
  const filterConditions = [];
  
  if (filterTerms.hasAttachments === true) {
    filterConditions.push('hasAttachments eq true');
  }
  
  if (filterTerms.unreadOnly === true) {
    filterConditions.push('isRead eq false');
  }
  
  // Add $filter parameter if we have any filter conditions
  if (filterConditions.length > 0) {
    params.$filter = filterConditions.join(' and ');
  }
}

/**
 * Add boolean filters as KQL terms for use with $search
 * Use this instead of addBooleanFilters when $search is present
 * @param {string[]} kqlTerms - Array of KQL terms to append to
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 */
function addBooleanFiltersAsKQL(kqlTerms, filterTerms) {
  if (filterTerms.hasAttachments === true) {
    kqlTerms.push('hasAttachments:true');
  }
  
  if (filterTerms.unreadOnly === true) {
    kqlTerms.push('isRead:false');
  }
}

/**
 * Format search results into a readable text format
 * @param {object} response - The API response object
 * @param {number} skip - Number of results skipped for display numbering
 * @returns {object} - MCP response object
 */
function formatSearchResults(response, skip = 0) {
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ 
        type: "text", 
        text: `No emails found matching your search criteria.`
      }]
    };
  }
  
  // Format results
  const emailList = response.value.map((email, index) => {
    const sender = email.from?.emailAddress || { name: 'Unknown', address: 'unknown' };
    const date = new Date(email.receivedDateTime).toLocaleString();
    const readStatus = email.isRead ? '' : '[UNREAD] ';
    const displayIndex = skip + index + 1;

    return `${displayIndex}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
  }).join("\n");
  
  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails matching your search criteria:\n\n${emailList}`
    }]
  };
}

module.exports = handleSearchEmails;
