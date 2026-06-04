# Local Development Setup with clasp

This guide outlines the steps required for a developer to set up their local environment, edit the Google Apps Script files locally, and push/pull code to the Google Sheets project using **clasp** (Command Line Apps Script Projects).

---

## Prerequisites

Before starting, ensure you have:
1. **Node.js** installed (v16 or higher recommended). You can verify with `node -v`.
2. **Access permissions**: You must be added as an **Editor** to the target Google Sheet containing the Apps Script project.

---

## Setup Steps

### 1. Clone the Repository
Clone the project repository to your laptop and navigate to the project directory:
```bash
git clone https://github.com/gangli72-ca/service_scheduler_dev.git
cd service_scheduler_dev
```

### 2. Install clasp
Install the Google clasp tool globally via `npm`:
```bash
npm install -g @google/clasp
```
*(Alternatively, if you prefer not to install globally, you can install it as a dev dependency with `npm install --save-dev @google/clasp` and run commands prefixed with `npx clasp`).*

### 3. Enable the Google Apps Script API
Google requires you to explicitly allow command-line tools to manage your scripts:
1. Go to the [Google Apps Script User Settings](https://script.google.com/home/usersettings).
2. Find the **Google Apps Script API** toggle.
3. Switch it to **ON** (Enabled).

> [!IMPORTANT]
> If you skip this step, any deployment/sync command (`clasp push`, `clasp pull`, etc.) will fail with a credentials/permission error.

### 4. Authenticate clasp
Log in to your Google account from the terminal:
```bash
clasp login
```
This command will open your default web browser and ask for permission. **You must log in using the Google Account that has Editor access to the Google Sheet.**

Once authenticated, a credentials file will be saved locally on your machine at `~/.clasprc.json`.

### 5. Verify the Connection
Since the project's unique script ID is already stored in `.clasp.json`, clasp will automatically know which remote Google Sheet to connect to. 

Verify the connection by running:
```bash
clasp status
```
This should show the files tracked by the script.

### 6. Pull the Latest Cloud Code
To make sure your local workspace is completely synchronized with the remote Apps Script state before editing, run:
```bash
clasp pull
```

---

## Daily Development Workflow

Once set up, you can edit the `.js` files using your preferred IDE (e.g., VS Code). Use the following commands to sync your changes:

* **Push local edits to Google Sheets:**
  ```bash
  clasp push
  ```
* **Auto-push on file save (Live Watch):**
  ```bash
  clasp push --watch
  ```
* **Pull remote edits to your laptop:**
  ```bash
  clasp pull
  ```
* **Open the remote Apps Script editor in your browser:**
  ```bash
  clasp open
  ```

---

## Troubleshooting

### "User has not enabled the Apps Script API"
If you see an error stating the API is not enabled:
- Double check that you enabled it in [Step 3](#3-enable-the-google-apps-script-api) for the **exact** Google Account you used to log in.
- Try logging out (`clasp logout`) and logging in again.

### "Permission denied" or "Could not find script"
- Ensure that your Google account has been granted **Editor** permissions to the target Google Sheet.
- Verify that the `scriptId` in `.clasp.json` matches the Script ID in the Google Sheet's Apps Script settings (**Project Settings** > **Script ID**).
