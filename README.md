# SVCA Children’s Ministry Service Scheduler

A complete **volunteer service scheduling automation system** built on **Google Sheets** and **Google Apps Script** for the SVCA Children’s Ministry.

This project manages volunteer availability, generates quarterly schedules with real-world constraints, sends notifications, and logs changes for auditing.

---

## 🚀 Features

### Core Capabilities
- Collect volunteer blackout (unavailable) dates
- Auto-mark EM (English Ministry) member blackout dates on special Sundays
- Automatically generate **quarterly service schedules**
- Two-step schedule creation: blank schedule with dropdowns, then auto-populate
- Enforce scheduling constraints:
  - One role per person per Sunday (except *floating roles*)
  - Couples cannot serve on the same Sunday
  - No volunteer serves 3 consecutive Sundays
- Highlight scheduling conflicts visually (auto-restoring after 15 seconds)
- Send email notifications to volunteers with confirm/decline links
- Web App for volunteers to confirm or decline assignments via email links
- Track schedule history and system actions
- Installable onEdit triggers for audit logging and access control

---

## 🧱 Architecture

**Platform**
- Google Sheets (data storage & UI)
- Google Apps Script (business logic & automation)

**Sheets Used**
- Roles
- Blackout Dates
- Schedule
- Schedule History
- Logs
- Config
- Couples
- Parent Helper

---

## 📊 Data Model

### Roles Sheet
| Column | Purpose |
|------|--------|
| A | Volunteer Name |
| B – (N-1) | Role eligibility (checkboxes), one column per role |
| Last column | Email address (header: "SVCA Email") |

> The email column is identified by the header name "SVCA Email", not a fixed column letter. Role columns span from B to the column before the email column.
>
> When a role checkbox is toggled, an installable trigger (`handleRolesEdit`) automatically updates the Schedule sheet dropdowns: adding or removing the volunteer from the relevant role's dropdown lists.

---

### Blackout Dates Sheet
- Generated quarterly
- Columns: `Name | Sunday Dates…`
- Volunteers mark unavailable dates using checkboxes
- Row-level editing restricted by logged-in email

---

### Schedule Sheet
- Column A: Date
- Columns B+: Service roles
- Auto-generated quarterly
- Dropdowns for manual adjustments
- Conflict highlighting:
  - Duplicate assignments
  - Consecutive Sundays
  - Couples serving together

---

### Schedule History
- Stores quarterly snapshots
- Old quarter data is removed before saving new schedules

---

### Logs
Records timestamped actions:
- Manual edits
- Email sends
- System operations (e.g., copying to history)

---

### Config
| Column | Cell(s) | Purpose |
|--------|---------|--------|
| A | A2 | Quarter start month (1–12, defaults to January) |
| B | B2 | Floating roles list (comma-separated) |
| C | C2 | Admin email addresses (comma-separated) |
| D | D2, D3, … | EM (English Ministry) member names (one per cell) |
| E | E2, E3, … | Combined/special Sunday dates (one date per cell) |
| F | F2, F3, … | Role names for lead lookup (paired with column G) |
| G | G2, G3, … | Lead email for the role in column F (for decline notifications) |
| H | H2, H3, … | Double-week roles (one per cell, read until empty) |

---

### Couples
- Two-column (Husband, Wife) mapping of couples
- Prevents spouses from serving on the same Sunday

---

### Parent Helper
- Column A: English name
- Column C: Chinese name
- Column D: Email
- Names on the Schedule appear as "<Chinese_name> <English_name>" (or just English name if no Chinese name)
- Used for email notifications and name matching on the Schedule sheet

---

## ⚙️ Major Script Features

### Quarter Calculation
- Dynamically computes quarter start/end
- Finds all Sundays within the quarter

---

### Blackout Date Management
- Generates blackout checkboxes
- Auto-marks blackout dates for EM members on combined/special Sundays (from Config columns D & E)
- Locks/unlocks sheet for volunteer input
- Installable trigger (`handleBlackoutEdit`) restricts edits to each volunteer's own row (matched by email), with admin bypass

---

### Schedule Generation
Schedule creation is a **two-step process**:

**Step 1: Create Blank Schedule** (`createBlankSchedule`)
- Appends new quarter's Sundays to the existing Schedule sheet
- Populates per-cell dropdown lists based on role eligibility and blackout dates
- Allows admin to manually pre-schedule specific cells before auto-populating

**Step 2: Auto Populate Schedule** (`autoPopulateSchedule`)
- Fills remaining empty cells using round-robin assignment per role
- Preserves manually pre-scheduled cells (admin pre-scheduling)
- Pre-scheduled volunteers are excluded from all auto-scheduling
- On combined dates (from Config column E), assigns "大堂 Combine" to "Lion Teacher"
- **Double-week roles** (from Config column H): same volunteer serves 2 consecutive Sundays
- Cross-quarter boundary: carries over incomplete double-week pairs from previous quarter
- Schedules double-week roles first, then remaining (non-double-week) roles
- "Parent Helper" roles are auto-set to "NA" (not auto-assigned)
- Enforces:
  - Role eligibility
  - Blackout compliance
  - One role per person/day (except floating roles)
  - No spouse conflicts
  - No back-to-back Sundays
  - No volunteer serves 3 consecutive weeks

---

### Conflict Highlighting
Detects and color-codes:
- **Light red** (#FFCCCC): Same person assigned to multiple roles on the same Sunday
- **Light yellow** (#FFF2CC): Same person serving 3 consecutive Sundays
- **Light blue** (#CCE5FF): Husband and wife serving on the same Sunday

Highlights are displayed for 15 seconds with a color legend, then automatically restored to the original cell backgrounds.

---

### Highlight One Person
- Select a volunteer's cell on the Schedule sheet to highlight all occurrences of that person (and their spouse) in light pink
- Highlights auto-restore after 10 seconds

---

### Email Notifications
- **Blackout notification emails**: Sent to selected volunteers on the Roles sheet; includes link to the Blackout Dates sheet (bilingual Chinese/English)
- **Upcoming Sunday assignment emails**: Sent to all volunteers assigned on the next upcoming Sunday; includes confirm/decline buttons per role assignment
- Email sending is logged automatically
- Supports time-driven triggers (gracefully falls back to logging when no UI is available)

---

## 🧭 Custom Menu

Adds a **Service Scheduler** menu to Google Sheets (visible to Editors only):

- Refresh Blackout Dates
- Lock Blackout Dates
- Unlock Blackout Dates
- *(separator)*
- Create Blank Schedule
- Auto Populate Schedule
- Highlight Conflicts
- Highlight One Person
- *(separator)*
- Copy to Schedule History
- *(separator)*
- Send Blackout Notification Emails
- Send Upcoming Sunday Emails

Non-editor users see an empty menu (no menu items).

---

## 🌐 Web App

The project includes a fully implemented Web App (`WebApp.js`) deployed as a Google Apps Script Web App:

- **Confirm**: Volunteer clicks the confirm link in their email → cell is highlighted light green (#90EE90) on the Schedule sheet
- **Decline**: Volunteer clicks the decline link → cell is highlighted light pink (#FFB6C1), and the role's lead (from Config columns F/G) is notified via email
- Validates that the Schedule cell still contains the expected volunteer name before processing
- Returns a styled HTML response page (green for confirm, yellow/orange for decline, red for errors)
- All actions are logged via `logAction()`

> The `WEB_APP_URL` is stored in **Script Properties** (per-project), allowing the same codebase to be pushed to both dev and prod environments without modification.

---

## 🔧 Installable Triggers

The following functions must be installed as **installable onEdit triggers** (not simple triggers) to work correctly:

| Function | Sheet | Purpose |
|----------|-------|---------|
| `handleScheduleEdit` | Schedule | Logs manual edits; highlights upcoming Sunday changes in green |
| `handleBlackoutEdit` | Blackout Dates | Enforces row-level access by email; admins can edit any row |
| `handleRolesEdit` | Roles | Syncs role checkbox changes to Schedule dropdowns; removes/adds volunteers |

---

## 📁 File Structure

| File | Purpose |
|------|---------|
| `Menus.js` | Custom menu setup (`onOpen`) |
| `BlackoutDates.js` | Blackout date generation, locking, EM marking, row-level access |
| `SundayScript.js` | Blank schedule creation, auto-population, conflict highlighting, history, triggers |
| `Email.js` | Blackout notification and upcoming Sunday assignment emails |
| `WebApp.js` | Web App for confirm/decline responses |
| `Utils.js` | Shared utilities: quarter calculation, logging, config readers, couples map |
| `appsscript.json` | Apps Script manifest (timezone, scopes, web app config) |

---

## ✅ Summary

This project is a **production-ready church volunteer scheduling platform** that:

- Reads structured volunteer data
- Collects availability automatically
- Builds intelligent schedules with real-world constraints
- Notifies volunteers reliably with confirm/decline workflow
- Maintains full audit history

---

## 📄 License

Internal use for SVCA Children’s Ministry.  
(Adjust licensing if open-sourcing.)

---

## 🤝 Contributions

Contributions are welcome from authorized collaborators.  
Please follow existing sheet structure and naming conventions.

---

## 📬 Questions or Enhancements

Feel free to open an issue or reach out to the project maintainer.
