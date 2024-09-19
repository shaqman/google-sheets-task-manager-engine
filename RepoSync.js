// List of labels to populate
const labels = [
  'Backlog',
  'To Do',
  'Blocked',
  'In Progress',
  'Code Review',
  'In Review',
  'Ready to Merge',
  'Merged',
  'Ready to Test',
  'Testing',
  'Testing Notes',
  'Ready for Deployment',
  'Ready to Implement',
  'Deployed',
  'On Hold',
  'Done'
];

/**
 * Retrieves the GitLab access token for a given base URL from Script Properties.
 * @param {string} baseUrl - The base URL of the GitLab instance.
 * @returns {string} - The corresponding GitLab access token.
 * @throws Will throw an error if no access token is found for the provided base URL.
 */
function getAccessTokenForBaseUrl(baseUrl) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Retrieve the access token from Script Properties
  const accessToken = scriptProperties.getProperty(baseUrl);
  
  if (!accessToken) {
    throw new Error(`No access token found for base URL: ${baseUrl}`);
  }
  
  return accessToken;
}

/**
 * Updates the labels of a GitLab issue by removing specified labels and adding a new label.
 * @param {string} issueUrl - The URL of the GitLab issue.
 * @param {string} newLabel - The new label to add to the issue.
 */
function updateGitLabIssueLabel(issueUrl, newLabel) {
  try {
    // Prepare the labels to remove, excluding newLabel
    const labelsToRemoveFiltered = labels.filter(label => label !== newLabel);
    const labelsToRemoveString = labelsToRemoveFiltered.join(',');

    // Prepare the payload
    const payload = {
      remove_labels: labelsToRemoveString,
      add_labels: newLabel
    };

    // Update the issue with the labels to remove and the new label to add
    const updatedIssue = gitLabApiRequest(
      issueUrl,
      'PUT',
      payload,
      'issues'
    );

    if (updatedIssue) {
      Logger.log(`Issue labels updated successfully.`);
    } else {
      Logger.log(`Failed to update issue.`);
    }
  } catch (error) {
    Logger.log('Error updating GitLab issue:', error);
  }
}


/**
 * Parses a GitLab issue URL to extract the base URL, project path, and issue IID.
 * @param {string} issueUrl - The URL of the GitLab issue.
 * @returns {Object} An object containing baseUrl, projectPath, and issueIID.
 */
function parseGitLabIssueUrl(issueUrl) {
  const urlPattern = /^(https?:\/\/[^/]+)\/([^/]+\/[^/]+(?:\/[^/]+)*)\/-\/issues\/(\d+)(\/.*)?$/;
  const match = issueUrl.match(urlPattern);

  if (!match) {
    throw new Error('Invalid GitLab issue URL.');
  }

  const baseUrl = match[1];
  const projectPath = match[2];
  const issueIID = match[3];

  Logger.log(baseUrl)
  Logger.log(projectPath)
  Logger.log(issueIID)

  return { baseUrl, projectPath, issueIID };
}

/**
 * Makes a request to the GitLab API.
 * @param {string} issueUrl - The URL of the GitLab issue.
 * @param {string} method - The HTTP method (GET, PUT, POST, DELETE).
 * @param {Object} [payload] - An optional payload for POST and PUT requests.
 * @param {string} apiType - The type of API endpoint ('issues', 'projects', etc.).
 * @returns {Object|null} The parsed JSON response from the API or null if failed.
 */
function gitLabApiRequest(issueUrl, method, payload, apiType = "issues") {
  try {
    // Parse the issue URL to extract the base URL, project path, and issue IID
    const { baseUrl, projectPath, issueIID } = parseGitLabIssueUrl(issueUrl);

    // Get the access token for the base URL
    const accessToken = getAccessTokenForBaseUrl(baseUrl);

    // Construct the API endpoint
    let endpoint = '';
    if (apiType === 'issues') {
      // For issue-related endpoints
      endpoint = `/projects/${encodeURIComponent(projectPath)}/issues/${issueIID}`;
    } else if (apiType === 'projects') {
      // For project-related endpoints
      endpoint = `/projects/${encodeURIComponent(projectPath)}`;
    } else {
      throw new Error(`Unknown API type: ${apiType}`);
    }

    const url = `${baseUrl}/api/v4${endpoint}`;

    const options = {
      method: method,
      headers: {
        'PRIVATE-TOKEN': accessToken,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    if (payload) {
      options.payload = JSON.stringify(payload);
      Logger.log(`Options Payload: ${options.payload}`);
    }

    // Clone options for logging without exposing the access token
    const optionsForLogging = { ...options };
    optionsForLogging.headers = { ...options.headers };
    optionsForLogging.headers['PRIVATE-TOKEN'] = 'REDACTED';

    Logger.log(`Request URL: ${url}`);
    Logger.log(`Options: ${JSON.stringify(optionsForLogging)}`);

    const response = UrlFetchApp.fetch(url, options);
    Logger.log(`Response Code: ${response.getResponseCode()}`);
    Logger.log(`Response Content: ${response.getContentText()}`);

    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
      return JSON.parse(response.getContentText());
    } else {
      Logger.log(
        `GitLab API request failed with status ${response.getResponseCode()}: ${response.getContentText()}`
      );
      return null;
    }
  } catch (error) {
    Logger.log('Error during UrlFetchApp.fetch:');
    Logger.log(`Message: ${error.message}`);
    Logger.log(`Stack Trace: ${error.stack}`);
    return null;
  }
}


/**
 * Retrieves the labels of a GitLab issue and returns the first matching label from the bottom of the array.
 * @param {string} issueUrl - The URL of the GitLab issue.
 * @returns {string|null} - The first matching label or null if none found.
 */
function getLatestMatchingLabelFromIssue(issueUrl) {
  try {
    // Parse the issue URL to extract project path and issue IID
    const { baseUrl, projectPath, issueIID } = parseGitLabIssueUrl(issueUrl);

    // URL-encode the project path
    const encodedProjectPath = encodeURIComponent(projectPath);

    // Get the issue details from GitLab API
    const issue = gitLabApiRequest(
      issueUrl,
      'GET',
      null
    );

    if (!issue) {
      Logger.log(`Issue #${issueIID} not found.`);
      return null;
    }

    const issueLabels = issue.labels;

    // Iterate over the issue labels from the bottom upwards
    for (let i = issueLabels.length - 1; i >= 0; i--) {
      const label = issueLabels[i];
      if (labels.includes(label)) {
        // Found a matching label
        return label;
      }
    }

    // No matching label found
    return null;

  } catch (error) {
    Logger.log('Error retrieving labels from GitLab issue:', error);
    return null;
  }
}