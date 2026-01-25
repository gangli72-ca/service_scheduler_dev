/**
 * Calculates the date range for the next quarter based on a configurable start month.
 *
 * It reads the “Quarter Start Month” from Config!A2 (defaults to January if empty or invalid), 
 * determines which quarter today falls into relative to that start month, then computes the 
 * first day (startDate) and last day (endDate) of the *following* quarter.
 *
 * @return {{startDate: Date, endDate: Date}} 
 *   - startDate: JavaScript Date for the first day of next quarter  
 *   - endDate:   JavaScript Date for the last day of next quarter
 */
function getNextQuarterRange() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    var today = new Date();
    var currentMonth = today.getMonth() + 1; // 1–12
    var year = today.getFullYear();

    // Read configured Q1 start month from Config!A2
    var raw = ss.getSheetByName("Config").getRange("A2").getValue();
    var q1Start = parseInt(raw, 10);
    if (isNaN(q1Start) || q1Start < 1 || q1Start > 12) {
        q1Start = 1;  // default to January
    }

    // Determine current quarter index (0–3) relative to q1Start
    var offset = (currentMonth - q1Start + 12) % 12;
    var currentQ = Math.floor(offset / 3);
    var nextQ = (currentQ + 1) % 4;

    // Compute start month/year of next quarter
    var startMonthRaw = (q1Start - 1) + nextQ * 3;
    var startYear = year + Math.floor(startMonthRaw / 12);
    var startMonth = (startMonthRaw % 12) + 1;
    var startDate = new Date(startYear, startMonth - 1, 1);

    // Compute end date of next quarter (last day of the third month)
    var endMonthRaw = startMonthRaw + 2;
    var endYear = year + Math.floor(endMonthRaw / 12);
    var endMonth = (endMonthRaw % 12) + 1;
    var endDate = new Date(endYear, endMonth, 0);

    return { startDate: startDate, endDate: endDate };
}


/**
 * Returns an array of all Sundays (as Date objects) in the next quarter.
 * @return {Date[]} Array of Sunday Date objects.
 */
function getSundaysForNextQuarter() {
    var quarterRange = getNextQuarterRange();
    var startDate = quarterRange.startDate;
    var endDate = quarterRange.endDate;

    // Find the first Sunday on or after the start date.
    var firstSunday = new Date(startDate);
    while (firstSunday.getDay() !== 0) { // 0 means Sunday
        firstSunday.setDate(firstSunday.getDate() + 1);
    }

    var sundays = [];
    for (var d = new Date(firstSunday); d <= endDate; d.setDate(d.getDate() + 7)) {
        sundays.push(new Date(d));
    }
    return sundays;
}

/**
 * Inserts a log entry into the "Logs" sheet.
 * The new log entry is inserted as the second row (right after the header).
 * Column A is the current timestamp in "yyyy/MM/dd HH:mm" format (Pacific Timezone),
 * Column B is the provided description,
 * and Column C is set to the logged in user's name extracted from their email.
 *
 * @param {string} description - The description text for the log entry.
 */
function logAction(description) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Logs");

    if (!logSheet) {
        // Create the Logs sheet if it doesn't exist and set the header row.
        logSheet = ss.insertSheet("Logs");
        logSheet.getRange(1, 1, 1, 3).setValues([["Date", "Description", "Person"]]);
    }

    // Insert a new row right before the first date row (row 2).
    logSheet.insertRowBefore(2);

    // Get the current time and format it in the Pacific Timezone.
    var now = new Date();
    var pacificTimeZone = "America/Los_Angeles";
    var formattedTimestamp = Utilities.formatDate(now, pacificTimeZone, "yyyy/MM/dd HH:mm");

    // Get the logged in user's email and extract the name before the "@".
    var email = Session.getActiveUser().getEmail();
    var name = email.split('@')[0];

    // Populate the new row: Column A = timestamp, Column B = description, Column C = name.
    logSheet.getRange(2, 1).setValue(formattedTimestamp);
    logSheet.getRange(2, 2).setValue(description);
    logSheet.getRange(2, 3).setValue(name);
}

/**
 * Reads the “floating roles” from Config!B2 (comma-separated),
 * trims each entry, and returns a clean array.
 *
 * @return {string[]} Array of floating role names.
 */
function getFloatingRoles() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = ss.getSheetByName('Config');
    var raw = config.getRange('B2').getDisplayValue();  // e.g. "Role 1,Role 2,Role 3"

    if (!raw) return [];

    return raw
        .split(',')                         // [ "Role 1", "Role 2", "Role 3" ]
        .map(function (item) {               // trim whitespace
            return item.trim();
        })
        .filter(function (item) {            // drop any empty strings
            return item.length > 0;
        });
}

/**
 * Builds a map from volunteer name -> email using the "Roles" sheet.
 * Assumes:
 *   - Column A: Name
 *   - Last column: Email
 *
 * @return {Object<string,string>} e.g. { "Alice": "alice@example.com", ... }
 */
function getVolunteerEmailMap() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rolesSheet = ss.getSheetByName("Roles");
    if (!rolesSheet) return {};

    var lastRow = rolesSheet.getLastRow();
    var lastCol = rolesSheet.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return {};

    // Grab all data including the last column (email)
    var range = rolesSheet.getRange(2, 1, lastRow - 1, lastCol);
    var data = range.getValues();

    var map = {};
    data.forEach(function (row) {
        var name = (row[0] || "").toString().trim();            // Column A
        var email = (row[lastCol - 1] || "").toString().trim();  // Last column = email
        if (name && email) {
            map[name] = email.toLowerCase();
        }
    });

    return map;
}

/**
 * Reads admin emails from Config!C2 (comma-separated) and returns
 * a normalized lowercase array.
 *
 * Example: "admin1@x.com, admin2@x.com"
 *   => ["admin1@x.com", "admin2@x.com"]
 *
 * @return {string[]} admin email list in lowercase.
 */
function getAdminEmails() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = ss.getSheetByName('Config');
    if (!config) return [];

    var raw = config.getRange('C2').getDisplayValue();
    if (!raw) return [];

    return raw
        .split(',')
        .map(function (item) {
            return item.trim().toLowerCase();
        })
        .filter(function (item) {
            return item.length > 0;
        });
}

/**
 * Returns the first scheduled Sunday strictly after today (excluding today) from the Schedule sheet.
 * If nothing is found, returns null.
 */
function findUpcomingSundayDate_(scheduleSheet) {
    var values = scheduleSheet.getDataRange().getValues();
    if (values.length < 2) {
        return null; // header only or empty
    }

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var upcoming = null;

    for (var i = 1; i < values.length; i++) {
        var d = values[i][0]; // column A is Date
        if (!(d instanceof Date)) {
            continue;
        }
        var dClean = new Date(d.getTime());
        dClean.setHours(0, 0, 0, 0);

        // Only look at dates strictly after today (excluding today)
        if (dClean > today) {
            if (upcoming === null || dClean < upcoming) {
                upcoming = dClean;
            }
        }
    }

    return upcoming;
}

/**
 * Builds a map of name -> spouse name from the "Couples" sheet.
 *
 * Couples sheet format:
 *   Row 1: Husband | Wife
 *   Rows 2+: pairs of names
 *
 * Example result:
 *   { "John": "Mary", "Mary": "John", ... }
 *
 * @return {Object<string,string>}
 */
function getCouplesMap() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Couples");
    if (!sheet) return {};

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    // Read columns A (Husband) and B (Wife) starting from row 2.
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var map = {};

    data.forEach(function (row) {
        var husband = (row[0] || "").toString().trim();
        var wife = (row[1] || "").toString().trim();

        if (husband && wife) {
            map[husband] = wife;
            map[wife] = husband;
        }
    });

    return map;
}

/**
 * Finds parent helper names on the Schedule sheet that cannot be found in either
 * the Roles sheet or the Parent Helper sheet.
 *
 * - Roles sheet: searches Column A for names.
 * - Parent Helper sheet: combines Column C (Chinese name) and Column A (English name)
 *   as "<Chinese_name> <English_Name>" or just "<English_name>" if no Chinese name.
 * - Schedule sheet: searches columns whose header contains "Helper".
 *
 * @return {string[]} Array of names that cannot be found in either sheet.
 */
function findUnmatchedParentHelpers() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Build a set of known names from the Roles sheet (Column A)
    var knownNames = {};
    var rolesSheet = ss.getSheetByName("Roles");
    if (rolesSheet && rolesSheet.getLastRow() >= 2) {
        var rolesData = rolesSheet.getRange(2, 1, rolesSheet.getLastRow() - 1, 1).getValues();
        rolesData.forEach(function (row) {
            var name = (row[0] || "").toString().trim();
            if (name) {
                knownNames[name.toLowerCase()] = true;
            }
        });
    }

    // 2. Add names from the Parent Helper sheet (combine Col C + Col A)
    var parentHelperSheet = ss.getSheetByName("Parent Helper");
    if (parentHelperSheet && parentHelperSheet.getLastRow() >= 2) {
        var phData = parentHelperSheet.getRange(2, 1, parentHelperSheet.getLastRow() - 1, 3).getValues();
        phData.forEach(function (row) {
            var englishName = (row[0] || "").toString().trim();  // Col A
            var chineseName = (row[2] || "").toString().trim();  // Col C

            // Construct full name as it appears on Schedule: "<Chinese> <English>" or just "<English>"
            var fullName = chineseName ? (chineseName + " " + englishName) : englishName;
            if (fullName) {
                knownNames[fullName.toLowerCase()] = true;
            }
        });
    }

    // 3. Collect helper names from the Schedule sheet (columns with "Helper" in header)
    var scheduleSheet = ss.getSheetByName("Schedule");
    if (!scheduleSheet || scheduleSheet.getLastRow() < 2) {
        return [];
    }

    var scheduleData = scheduleSheet.getDataRange().getValues();
    var headers = scheduleData[0];

    // Find column indices where header contains "Helper"
    var helperColIndices = [];
    for (var col = 0; col < headers.length; col++) {
        var header = (headers[col] || "").toString();
        if (header.indexOf("Helper") !== -1) {
            helperColIndices.push(col);
        }
    }

    // Collect all unique helper names from those columns
    var scheduleNames = {};
    for (var row = 1; row < scheduleData.length; row++) {
        helperColIndices.forEach(function (colIdx) {
            var name = (scheduleData[row][colIdx] || "").toString().trim();
            if (name) {
                scheduleNames[name] = true;
            }
        });
    }

    // 4. Find names that are NOT in knownNames
    var unmatched = [];
    Object.keys(scheduleNames).forEach(function (name) {
        if (!knownNames[name.toLowerCase()]) {
            unmatched.push(name);
        }
    });

    // Sort alphabetically for easier reading
    unmatched.sort();

    Logger.log("Unmatched parent helpers: " + JSON.stringify(unmatched));
    return unmatched;
}
