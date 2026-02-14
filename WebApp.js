/**
 * Web App for handling confirm/decline responses from weekly notification emails.
 * 
 * Deploy as a web app with:
 *   - Execute as: Your account
 *   - Who has access: Anyone with the link
 * 
 * After deployment, update WEB_APP_URL with your actual deployment URL.
 */

// IMPORTANT: Update this URL after deploying the web app
var WEB_APP_URL = "https://script.google.com/a/macros/svca.cc/s/AKfycbxv2YnLAJcVOY1g5EB2LtAKHltvqeAc2JibniaqXpUs4MqZlG4MTw1AWs5qkM8je_y6/exec";

/**
 * Main entry point for the web app.
 * Handles GET requests with action, name, role, and date parameters.
 * 
 * @param {Object} e - Event object containing URL parameters.
 * @return {HtmlOutput} HTML response to display to the user.
 */
function doGet(e) {
    var params = e.parameter;
    var action = params.action;
    var name = params.name;
    var role = params.role;
    var date = params.date;

    // Validate required parameters
    if (!action || !name || !role || !date) {
        return createHtmlResponse(
            "Error",
            "Missing required parameters. Please use the link from your email.",
            false
        );
    }

    try {
        if (action === "confirm") {
            return handleConfirm(name, role, date);
        } else if (action === "decline") {
            return handleDecline(name, role, date);
        } else {
            return createHtmlResponse(
                "Error",
                "Invalid action. Please use the link from your email.",
                false
            );
        }
    } catch (error) {
        logAction("Web app error: " + error.message);
        return createHtmlResponse(
            "Error",
            "An error occurred: " + error.message,
            false
        );
    }
}

/**
 * Handles a confirm action.
 * Sets the Schedule cell background to light green.
 * 
 * @param {string} name - Volunteer name.
 * @param {string} role - Role name.
 * @param {string} date - Date string (format: MM/dd/yyyy).
 * @return {HtmlOutput} Success or error HTML response.
 */
function handleConfirm(name, role, date) {
    var cell = findScheduleCell(role, date);
    if (!cell) {
        return createHtmlResponse(
            "Error",
            "Could not find the schedule entry for " + role + " on " + date + ".",
            false
        );
    }

    // Verify the cell contains the expected volunteer name
    var cellValue = cell.getValue();
    if (cellValue !== name) {
        return createHtmlResponse(
            "Warning",
            "The schedule has been modified. Expected " + name + " but found " + cellValue + ".",
            false
        );
    }

    // Set background to light green
    cell.setBackground("#90EE90"); // Light green

    logAction(name + " confirmed assignment for **" + role + "** on " + date);

    return createHtmlResponse(
        "Confirmed!",
        "Thank you, " + name + "! Your assignment for <strong>" + role + "</strong> on " + date + " has been confirmed.",
        true
    );
}

/**
 * Handles a decline action.
 * Sets the Schedule cell background to pink and sends email to the lead.
 * 
 * @param {string} name - Volunteer name.
 * @param {string} role - Role name.
 * @param {string} date - Date string (format: MM/dd/yyyy).
 * @return {HtmlOutput} Success or error HTML response.
 */
function handleDecline(name, role, date) {
    var cell = findScheduleCell(role, date);
    if (!cell) {
        return createHtmlResponse(
            "Error",
            "Could not find the schedule entry for " + role + " on " + date + ".",
            false
        );
    }

    // Verify the cell contains the expected volunteer name
    var cellValue = cell.getValue();
    if (cellValue !== name) {
        return createHtmlResponse(
            "Warning",
            "The schedule has been modified. Expected " + name + " but found " + cellValue + ".",
            false
        );
    }

    // Set background to pink
    cell.setBackground("#FFB6C1"); // Light pink

    // Get lead email and send notification
    var leadEmail = getLeadEmail(role);
    if (leadEmail) {
        var subject = "SVCA Sunday School - Volunteer Declined: " + role + " on " + date;
        var body = "Hello,\n\n" +
            name + " has declined the assignment for " + role + " on " + date + ".\n\n" +
            "Please arrange for a replacement volunteer.\n\n" +
            "SVCA Children's Ministry Scheduler";

        MailApp.sendEmail({
            to: leadEmail,
            subject: subject,
            body: body
        });

        logAction(name + " declined assignment for **" + role + "** on " + date + ". Notified lead: " + leadEmail);
    } else {
        logAction(name + " declined assignment for **" + role + "** on " + date + ". (No lead email found)");
    }

    return createHtmlResponse(
        "Declined",
        "Thank you for letting us know, " + name + ". Your assignment for <strong>" + role + "</strong> on " + date + " has been marked as declined. The ministry lead has been notified.",
        "decline"
    );
}

/**
 * Finds the cell on the Schedule sheet for a given role and date.
 * 
 * @param {string} role - Role name (column header).
 * @param {string} dateStr - Date string (format: MM/dd/yyyy).
 * @return {Range|null} The cell Range, or null if not found.
 */
function findScheduleCell(role, dateStr) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var scheduleSheet = ss.getSheetByName("Schedule");
    if (!scheduleSheet) return null;

    var tz = ss.getSpreadsheetTimeZone();
    var dateFormat = "MM/dd/yyyy";

    // Find the column index for this role
    var headers = scheduleSheet.getRange(1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
    var roleColIndex = -1;
    for (var c = 0; c < headers.length; c++) {
        if (headers[c] === role) {
            roleColIndex = c + 1; // 1-indexed
            break;
        }
    }
    if (roleColIndex < 2) return null; // Role not found or it's the Date column

    // Find the row index for this date
    var dateColumn = scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow() - 1, 1).getValues();
    var targetRowIndex = -1;
    for (var r = 0; r < dateColumn.length; r++) {
        var cellDate = dateColumn[r][0];
        if (cellDate instanceof Date) {
            var formattedDate = Utilities.formatDate(cellDate, tz, dateFormat);
            if (formattedDate === dateStr) {
                targetRowIndex = r + 2; // +2 because data starts at row 2 and r is 0-indexed
                break;
            }
        }
    }
    if (targetRowIndex < 2) return null; // Date not found

    return scheduleSheet.getRange(targetRowIndex, roleColIndex);
}

/**
 * Gets the lead email for a given role from the Config sheet.
 * Reads from Config columns F (Role) and G (Lead Email), starting from row 2.
 * 
 * @param {string} role - Role name to look up.
 * @return {string|null} Lead email address, or null if not found.
 */
function getLeadEmail(role) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) return null;

    var lastRow = configSheet.getLastRow();
    if (lastRow < 2) return null;

    // Read columns F and G starting from row 2
    var data = configSheet.getRange(2, 6, lastRow - 1, 2).getValues(); // F=6, G=7

    for (var i = 0; i < data.length; i++) {
        var configRole = (data[i][0] || "").toString().trim();
        var leadEmail = (data[i][1] || "").toString().trim();

        if (configRole === role && leadEmail) {
            return leadEmail;
        }
    }

    return null;
}

/**
 * Creates an HTML response page.
 * 
 * @param {string} title - Page title.
 * @param {string} message - Message to display.
 * @param {boolean|string} status - true for success (green), 'decline' for decline (pink), false for error (red).
 * @return {HtmlOutput} Formatted HTML response.
 */
function createHtmlResponse(title, message, status) {
    var bgColor, textColor, borderColor, icon;

    if (status === true) {
        // Confirm success - green
        bgColor = "#d4edda";
        textColor = "#155724";
        borderColor = "#c3e6cb";
        icon = "✅";
    } else if (status === "decline") {
        // Decline acknowledgment - pink/orange
        bgColor = "#fff3cd";
        textColor = "#856404";
        borderColor = "#ffeeba";
        icon = "📋";
    } else {
        // Error - red
        bgColor = "#f8d7da";
        textColor = "#721c24";
        borderColor = "#f5c6cb";
        icon = "⚠️";
    }

    var html = '<!DOCTYPE html>' +
        '<html>' +
        '<head>' +
        '<meta charset="UTF-8">' +
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
        '<title>' + title + ' - SVCA Children\'s Ministry</title>' +
        '<style>' +
        'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }' +
        '.container { max-width: 600px; margin: 50px auto; padding: 30px; background-color: ' + bgColor + '; border: 1px solid ' + borderColor + '; border-radius: 8px; text-align: center; }' +
        'h1 { color: ' + textColor + '; margin-bottom: 20px; }' +
        'p { color: ' + textColor + '; font-size: 16px; line-height: 1.5; }' +
        '.logo { font-size: 48px; margin-bottom: 20px; }' +
        '</style>' +
        '</head>' +
        '<body>' +
        '<div class="container">' +
        '<div class="logo">' + icon + '</div>' +
        '<h1>' + title + '</h1>' +
        '<p>' + message + '</p>' +
        '</div>' +
        '</body>' +
        '</html>';

    return HtmlService.createHtmlOutput(html)
        .setTitle(title + " - SVCA Children's Ministry");
}
