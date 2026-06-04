# Google Apps Script Project SOP

This document outlines the standard operating procedures for managing the Service Scheduler project using `clasp` (Command Line Apps Script Projects) and `git`.

## Prerequisites

- **Node.js & npm**: Installed.
- **Clasp**: Installed globally (`npm install -g @google/clasp`).
- **Git**: Installed and initialized.
- **Logged in**: Run `clasp login` if you haven't recently.

## Daily Workflow

### 1. Start of Session: Pull Latest Changes
Before starting work, ensuring your local environment matches the remote Google Apps Script project (useful if someone edited the code in the browser).

```bash
# Pull changes from Google Drive to local folder
clasp pull
```
*Note: If you have local uncommitted changes, commit or stash them first to avoid overwriting.*

### 2. Development: Edit & Test
Make your code changes in your local IDE (VS Code). 

### 3. Version Control: Commit Changes
Once you are satisfied with a chunk of work, save it to Git.

```bash
# Check status of modified files
git status

# Add files to staging
git add .

# Commit with a descriptive message
git commit -m "feat: Add new email notification logic"
```

### 4. Deployment: Push to Apps Script
Push your local code to the Google Apps Script server.

```bash
# Push local files to Google Drive
clasp push
```
*Note: `clasp push` overwrites the code on the server. Always ensure you pulled first if others are working on the project.*

### 5. (Optional) Watch Mode
If you want to push changes automatically every time you save a file:
```bash
clasp push --watch
```

## Handling Conflicts

If `clasp push` fails because the remote project has changed (Manifest errors or version mismatch):

1.  **Pull remote changes**: `clasp pull`
2.  **Resolve conflicts**: Git won't help with `clasp pull` conflicts directly as clasp overwrites. 
    *   *Best Practice*: If you know the remote has changes, backup your local work, run `clasp pull`, then re-apply your changes.
    *   *Alternative*: Force push (Be careful!): `clasp push --force` (This wipes out remote changes).

## Reference Commands

| Command | Description |
| :--- | :--- |
| `clasp login` | Login to Google account |
| `clasp pull` | Fetch code from Google Apps Script |
| `clasp push` | Upload code to Google Apps Script |
| `clasp open` | Open the project in the browser |
| `git status` | Show modified files |
| `git log` | View history of changes |
