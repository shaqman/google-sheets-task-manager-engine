# Google Sheets Task Manager Engine

## Summary

This project is a task manager engine built on Google Sheets and Google Apps Script. It automates task management functions such as scheduling, monitoring, and synchronizing data with repositories. The system integrates seamlessly with Google Sheets to manage tasks efficiently.

## Features

- **Initialization:** Sets up the necessary configurations and initializes the task manager.
- **Monitoring:** Tracks the progress of tasks and provides updates on their status.
- **Repository Synchronization:** Syncs data between Google Sheets and external repositories.
- **Scheduling:** Automates the scheduling of tasks, allowing for efficient task management.
- **Utility Functions:** Provides helper functions to support the main features of the task manager.

## Getting Started

### Prerequisites:
- Node.js installed on your machine.
- `clasp` installed globally via npm.
- A Google account with access to Google Apps Script.

### Installation:

1. **Install Clasp:**
   If you haven't installed `clasp` yet, you can do so by running the following command:
   ```
   npm install -g @google/clasp
   ```

2. **Login to Clasp:**
   Authenticate `clasp` with your Google account:
   ```
   clasp login
   ```

3. **Clone the Existing Project:**
   Clone this project using `clasp`. You will need the script's project ID, which is included in the Apps Script editor URL or can be found in the `appsscript.json` file.
   ```
   clasp clone <scriptId>
   ```
   Replace `<scriptId>` with the actual project ID of this script.

4. **Start Coding:**
   After cloning, the project files will be downloaded to your local environment. You can now start coding using your preferred text editor or IDE.

5. **Pushing Changes:**
   Once you've made changes to the code, you can push these changes back to Google Apps Script:
   ```
   clasp push
   ```

6. **Pulling Updates:**
   If there are updates made in the Google Apps Script editor that you need to pull into your local environment:
   ```
   clasp pull
   ```

## Optional: Adopting and Hosting the Script Yourself

If you wish to make a copy of this script and host it under your own Google account:

1. **Create a New Google Apps Script Project:**
   In the Google Apps Script dashboard, create a new project.

2. **Retrieve the New Script ID:**
   After creating the project, get the new `scriptId` from the URL or the project settings.

3. **Update the Project with the New Script ID:**
   Replace the existing `scriptId` in your local `clasp` project with the new one:
   ```
   clasp clone <newScriptId>
   ```
   Or, if you already have the code locally and want to push it to a new project:
   ```
   clasp setting scriptId <newScriptId>
   ```

4. **Push the Code to the New Project:**
   After updating the `scriptId`, push the local code to the new Apps Script project:
   ```
   clasp push
   ```

5. **Verify the New Project:**
   Go to the Google Apps Script dashboard and verify that your code has been successfully transferred to the new project.

## Additional Resources:
- Clasp Documentation: https://developers.google.com/apps-script/guides/clasp
- Google Apps Script Overview: https://developers.google.com/apps-script/overview

Happy coding!
