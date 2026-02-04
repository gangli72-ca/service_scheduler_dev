/** Hello
 * Refreshes the Blackout Dates sheet.
 * - Reads volunteer names from the "Roles" sheet.
 * - Populates the header row with Sunday dates for the next quarter.
 * - Fills the first column with volunteer names and inserts checkboxes for each Sunday.
 * - Applies background colors to header cells.
 */
function refreshBlackoutDates() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rolesSheet = ss.getSheetByName("Roles");
    var blackoutSheet = ss.getSheetByName("Blackout Dates");

    // Get volunteer names from the Roles sheet (assumes names are in column A starting at row 2).
    var lastRow = rolesSheet.getLastRow();
    if (lastRow < 2) {
        SpreadsheetApp.getUi().alert("No volunteer names found in the Roles sheet.");
        return;
    }
    var namesRange = rolesSheet.getRange(2, 1, lastRow - 1, 1);
    var namesData = namesRange.getValues();

    // Get all Sundays for the next quarter.
    var sundays = getSundaysForNextQuarter();

    // Thoroughly clear the Blackout Dates sheet, including old checkboxes/data validations
    var maxRows = blackoutSheet.getMaxRows();
    var maxCols = blackoutSheet.getMaxColumns();
    var fullRange = blackoutSheet.getRange(1, 1, maxRows, maxCols);
    fullRange.clearContent();         // remove all values
    fullRange.clearFormat();          // remove colors/borders/fonts
    fullRange.clearDataValidations(); // remove old checkbox rules

    // Set header row: first header is "Name"; subsequent headers are Sunday dates.
    var header = ["Name"];
    var dateFormat = "MM/dd/yyyy";
    sundays.forEach(function (date) {
        header.push(Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), dateFormat));
    });
    var headerRange = blackoutSheet.getRange(1, 1, 1, header.length);
    headerRange.setValues([header]);
    headerRange.setBackground("#CCCCCC"); // column header background color

    // Write volunteer names into column A starting at row 2.
    var nameRange = blackoutSheet.getRange(2, 1, namesData.length, 1);
    nameRange.setValues(namesData);
    nameRange.setBackground("#DDDDDD"); // row header background color

    // Fill remaining cells with checkboxes.
    var numRows = namesData.length;
    var numCols = header.length - 1;
    var dataRange = blackoutSheet.getRange(2, 2, numRows, numCols);
    dataRange.clearContent();

    // Set data validation for checkboxes.
    var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    dataRange.setDataValidation(rule);

    SpreadsheetApp.getUi().alert("Blackout Dates sheet refreshed successfully.");
}

/**
 * Locks the Blackout Dates sheet so that checkboxes (and all cells) become view-only.
 * Removes any previous protections on the sheet before applying a new one.
 */
function lockBlackoutDates() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var blackoutSheet = ss.getSheetByName("Blackout Dates");

    // Remove existing protections.
    var protections = blackoutSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(function (protection) {
        protection.remove();
    });

    // Protect the entire sheet.
    var protection = blackoutSheet.protect().setDescription("Blackout Dates Locked");
    protection.setWarningOnly(false);

    // Allow only the effective user (admin) to edit.
    var me = Session.getEffectiveUser();
    protection.addEditor(me);

    // Remove any other editors.
    var editors = protection.getEditors();
    editors.forEach(function (editor) {
        if (editor.getEmail() !== me.getEmail()) {
            protection.removeEditor(editor);
        }
    });

    // Ensure domain users cannot edit.
    if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
    }

    SpreadsheetApp.getUi().alert("Blackout Dates sheet has been locked.");
}

function unlockBlackoutDates() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var blackoutSheet = ss.getSheetByName("Blackout Dates");

    // Get all sheet protections on the Blackout Dates sheet.
    var protections = blackoutSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

    // Remove each protection.
    protections.forEach(function (protection) {
        protection.remove();
    });

    SpreadsheetApp.getUi().alert("Blackout Dates sheet unlocked successfully.");
}

/**
 * Installable onEdit trigger to enforce:
 * A user may only edit the blackout row that corresponds to their own email,
 * UNLESS they are an admin (email listed in Config!C2), in which case they
 * may edit any row.
 */
function handleBlackoutEdit(e) {
    if (!e) return;

    var range = e.range;
    var sheet = range.getSheet();

    // Only enforce on the "Blackout Dates" sheet
    if (sheet.getName() !== "Blackout Dates") return;

    var row = range.getRow();
    var col = range.getColumn();

    // Ignore header row and name column
    if (row < 2 || col < 2) return;

    // Get the current user email
    var userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
        // If we can't see the user email (e.g. some account types), safest is to block edits.
        if (typeof e.oldValue !== "undefined") {
            range.setValue(e.oldValue);
        } else {
            range.clearContent();
        }
        SpreadsheetApp.getActive().toast(
            "Edit not allowed: unable to verify your account email.",
            "Blackout Dates",
            5
        );
        return;
    }
    userEmail = userEmail.toLowerCase().trim();

    // --- Admin bypass: admins can edit any row ---
    var adminEmails = getAdminEmails();  // from Config!C2
    if (adminEmails.indexOf(userEmail) !== -1) {
        // Admin – allow the edit with no further checks
        return;
    }
    // ------------------------------------------------

    // The volunteer name for this row (col A)
    var volunteerName = sheet.getRange(row, 1).getDisplayValue().trim();
    if (!volunteerName) return;  // no name => nothing to enforce

    // Map name -> email from Roles sheet
    var emailMap = getVolunteerEmailMap();
    var expectedEmail = emailMap[volunteerName];

    if (!expectedEmail) {
        // No email configured for this name; block for regular users
        if (typeof e.oldValue !== "undefined") {
            range.setValue(e.oldValue);
        } else {
            range.clearContent();
        }
        SpreadsheetApp.getActive().toast(
            "Edit not allowed: no email configured for \"" + volunteerName + "\" in Roles.",
            "Blackout Dates",
            5
        );
        return;
    }

    expectedEmail = expectedEmail.toLowerCase().trim();

    // If the logged-in email doesn't match the row's email, revert
    if (expectedEmail !== userEmail) {
        if (typeof e.oldValue !== "undefined") {
            range.setValue(e.oldValue);
        } else {
            range.clearContent();
        }
        SpreadsheetApp.getActive().toast(
            "You can only edit your own blackout dates row.",
            "Blackout Dates",
            5
        );
        return;
    }

    // If we reach here, the user email matches the row's email → edit allowed
}

/**
 * Marks blackout dates for EM (English Ministry) members on special Sunday dates.
 * 
 * Reads from the Config sheet:
 *   - Column D (rows 2+): EM member names
 *   - Column E (rows 2+): Special Sunday dates
 * 
 * For each EM member and each special Sunday date, finds the corresponding
 * cell on the Blackout Dates sheet and checks the checkbox (sets value to true).
 */
function markEMBlackoutDates() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    var blackoutSheet = ss.getSheetByName("Blackout Dates");

    if (!configSheet || !blackoutSheet) {
        SpreadsheetApp.getUi().alert("Config or Blackout Dates sheet not found.");
        return;
    }

    // Read EM members from Config Column D (starting row 2)
    var configLastRow = configSheet.getLastRow();
    if (configLastRow < 2) {
        SpreadsheetApp.getUi().alert("No data found in Config sheet.");
        return;
    }

    // Get EM members (Column D)
    var emMembersRange = configSheet.getRange(2, 4, configLastRow - 1, 1); // D2:D
    var emMembersData = emMembersRange.getValues();
    var emMembers = emMembersData
        .map(function (row) { return (row[0] || "").toString().trim(); })
        .filter(function (name) { return name.length > 0; });

    // Get special Sunday dates (Column E)
    var datesRange = configSheet.getRange(2, 5, configLastRow - 1, 1); // E2:E
    var datesData = datesRange.getValues();
    var specialDates = datesData
        .filter(function (row) { return row[0] instanceof Date; })
        .map(function (row) { return row[0]; });

    if (emMembers.length === 0) {
        SpreadsheetApp.getUi().alert("No EM members found in Config Column D.");
        return;
    }

    if (specialDates.length === 0) {
        SpreadsheetApp.getUi().alert("No special Sunday dates found in Config Column E.");
        return;
    }

    // Build a map of volunteer names to their row indices on Blackout Dates sheet
    var blackoutLastRow = blackoutSheet.getLastRow();
    var blackoutLastCol = blackoutSheet.getLastColumn();

    if (blackoutLastRow < 2 || blackoutLastCol < 2) {
        SpreadsheetApp.getUi().alert("Blackout Dates sheet appears to be empty or not set up.");
        return;
    }

    // Get all volunteer names from Column A (row 2 onwards)
    var namesData = blackoutSheet.getRange(2, 1, blackoutLastRow - 1, 1).getValues();
    var nameToRow = {};
    for (var i = 0; i < namesData.length; i++) {
        var name = (namesData[i][0] || "").toString().trim();
        if (name) {
            nameToRow[name] = i + 2; // row index (1-indexed, starting from row 2)
        }
    }

    // Get all date headers from row 1 (column 2 onwards)
    // Use getDisplayValues() to get formatted date strings as shown in the sheet
    var headersData = blackoutSheet.getRange(1, 2, 1, blackoutLastCol - 1).getDisplayValues()[0];
    var dateFormat = "MM/dd/yyyy";
    var tz = ss.getSpreadsheetTimeZone();
    var dateToCol = {};
    for (var j = 0; j < headersData.length; j++) {
        var headerText = (headersData[j] || "").toString().trim();
        if (headerText) {
            dateToCol[headerText] = j + 2; // column index (1-indexed, starting from column 2)
        }
    }

    // Mark blackout dates for each EM member and special date
    var markedCount = 0;
    var notFoundMembers = [];
    var notFoundDates = [];

    emMembers.forEach(function (memberName) {
        var row = nameToRow[memberName];
        if (!row) {
            if (notFoundMembers.indexOf(memberName) === -1) {
                notFoundMembers.push(memberName);
            }
            return;
        }

        specialDates.forEach(function (date) {
            var formattedDate = Utilities.formatDate(date, tz, dateFormat);
            var col = dateToCol[formattedDate];
            if (!col) {
                if (notFoundDates.indexOf(formattedDate) === -1) {
                    notFoundDates.push(formattedDate);
                }
                return;
            }

            // Check the checkbox (set to true)
            blackoutSheet.getRange(row, col).setValue(true);
            markedCount++;
        });
    });

    // Build result message
    var message = "Marked " + markedCount + " blackout date(s) for EM members.";

    if (notFoundMembers.length > 0) {
        message += "\n\nEM members not found on Blackout Dates sheet:\n- " + notFoundMembers.join("\n- ");
    }

    if (notFoundDates.length > 0) {
        message += "\n\nDates not found on Blackout Dates sheet:\n- " + notFoundDates.join("\n- ");
    }

    SpreadsheetApp.getUi().alert(message);
}
