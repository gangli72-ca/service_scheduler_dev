# SVCA Children’s Ministry Service Scheduler

A complete **volunteer service scheduling automation system** built on **Google Sheets** and **Google Apps Script** for the SVCA Children’s Ministry.

This project manages volunteer availability, generates quarterly schedules with real-world constraints, sends notifications, and logs changes for auditing.

---

## 🚀 Features

### Core Capabilities
- Collect volunteer blackout (unavailable) dates
- Automatically generate **quarterly service schedules**
- Enforce scheduling constraints:
  - One role per person per Sunday (except *floating roles*)
  - Couples cannot serve on the same Sunday
  - No volunteer serves on consecutive Sundays
- Highlight scheduling conflicts visually
- Send email notifications to volunteers
- Track schedule history and system actions

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

---

## 📊 Data Model

### Roles Sheet
| Column | Purpose |
|------|--------|
| A | Volunteer Name |
| B–P | Role eligibility (checkboxes) |
| Q | Email address |

> Optional enhancements: phone numbers for SMS/WhatsApp notifications.

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
| Cell | Purpose |
|----|--------|
| A2 | Quarter start month |
| B2 | Floating roles list |
| C2 | Admin email addresses (comma-separated) |

---

### Couples
- Two-column (Husband, Wife) mapping of couples
- Prevents spouses from serving on the same Sunday

---

## ⚙️ Major Script Features

### Quarter Calculation
- Dynamically computes quarter start/end
- Finds all Sundays within the quarter

---

### Blackout Date Management
- Generates blackout checkboxes
- Locks/unlocks sheet for volunteer input
- Restricts edits to each volunteer’s own row

---

### Schedule Generation
- Builds a full quarterly schedule
- Enforces:
  - Role eligibility
  - Blackout compliance
  - One role per person/day (except floating roles)
  - No spouse conflicts
  - No consecutive Sundays
- Uses round-robin assignment per role
- Populates dropdowns for manual adjustments

---

### Conflict Highlighting
Detects:
- Duplicate assignments on the same Sunday
- Consecutive Sunday assignments
- Couples serving on the same day

Includes a **Clear Highlights** option.

---

### Email Notifications
- Blackout reminder emails
- Weekly upcoming Sunday assignment reminders
- Quarterly schedule notifications
- Email sending is logged automatically

---

## 🧭 Custom Menu

Adds a **Service Scheduler** menu to Google Sheets:

- Refresh Blackout Dates
- Lock / Unlock Blackout Dates
- Auto Populate Schedule
- Highlight Conflicts
- Clear Highlights
- Copy to Schedule History
- Send Blackout Reminder Email
- Send Weekly Assignment Reminder Email

Menu items can be hidden for non-editor users.

---

## 🌐 Web App (Optional / In Progress)

Planned features using Apps Script Web App endpoints:

- Confirmation links in email/SMS
- Volunteer responses:
  - Confirm assignment
  - Decline assignment
- Schedule updates
- Action logging

---

## ✅ Summary

This project is a **production-ready church volunteer scheduling platform** that:

- Reads structured volunteer data
- Collects availability automatically
- Builds intelligent schedules with real-world constraints
- Notifies volunteers reliably
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
