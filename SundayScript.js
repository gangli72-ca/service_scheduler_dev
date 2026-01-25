/**
 * Automatically populates the Schedule sheet with service dates and volunteer assignments.
 * - Populates column A with all Sundays in the next quarter.
 * - Uses volunteer role information from the Roles sheet and blackout info to assign volunteers.
 * - Applies round-robin assignment for each role in a persistent fashion, ensuring that if a volunteer is skipped due to a blackout, the rotation continues from that point.
 * - Ensures no volunteer is assigned more than one role on the same day.
 * - Sets dropdowns in each position column based on qualified volunteers.
 * - Applies header background colors.
 */
function autoPopulateSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("Schedule");
  var rolesSheet = ss.getSheetByName("Roles");
  var blackoutSheet = ss.getSheetByName("Blackout Dates");
  var dateFormat = "MM/dd/yyyy";

  // Get all Sundays for the next quarter.
  var sundays = getSundaysForNextQuarter();

  // Thoroughly clear the Schedule sheet, including any old data validations
  // that might be lingering in columns to the right of the active area.
  var maxRows = scheduleSheet.getMaxRows();
  var maxCols = scheduleSheet.getMaxColumns();
  var fullRange = scheduleSheet.getRange(1, 1, maxRows, maxCols);

  fullRange.clearContent();        // remove all values
  fullRange.clearFormat();         // remove background colors, fonts, etc.
  fullRange.clearDataValidations(); // remove ALL dropdowns / validation rules 

  // Get role headers from the Roles sheet.
  // Assumes: Col A = Name, Col B..(second last) = roles, LAST column = Email.
  var lastCol = rolesSheet.getLastColumn();
  var rolesHeader = rolesSheet.getRange(1, 2, 1, lastCol - 2).getValues()[0];

  // Write header row in Schedule sheet: Column A: "Date", columns B onward: role names.
  var scheduleHeader = ["Date"].concat(rolesHeader);
  var headerRange = scheduleSheet.getRange(1, 1, 1, scheduleHeader.length);
  headerRange.setValues([scheduleHeader]);
  headerRange.setBackground("#CCCCCC");

  // Write Sunday dates into column A (starting at row 2).
  var dateValues = sundays.map(function (date) {
    return [date];
  });
  scheduleSheet.getRange(2, 1, sundays.length, 1).setValues(dateValues);
  scheduleSheet.getRange(2, 1, sundays.length, 1).setNumberFormat(dateFormat);
  scheduleSheet.getRange(2, 1, sundays.length, 1).setBackground("#DDDDDD");

  // Build mapping of roles to qualified volunteers from the Roles sheet.
  // Assumes data starts at row 2: column A is volunteer name; columns B onward are checkboxes.
  var rolesDataRange = rolesSheet.getRange(2, 1, rolesSheet.getLastRow() - 1, rolesSheet.getLastColumn());
  var rolesData = rolesDataRange.getValues();
  var roleVolunteers = {};
  rolesHeader.forEach(function (role) {
    roleVolunteers[role] = [];
  });
  rolesData.forEach(function (row) {
    var name = row[0];
    rolesHeader.forEach(function (role, i) {
      if (row[i + 1] === true) {
        roleVolunteers[role].push(name);
      }
    });
  });

  // NOTE: We no longer set one validation rule per column here.
  // Instead, we will build per-cell dropdowns AFTER loading blackout data,
  // so that each date’s dropdown excludes volunteers who are blacked out on that date.

  // Load blackout data.
  var blackoutData = blackoutSheet.getDataRange().getValues();
  var blackoutHeader = blackoutData[0];
  var blackoutDateMap = {};
  for (var j = 1; j < blackoutHeader.length; j++) {
    var d = blackoutHeader[j];
    if (d instanceof Date) {
      var formatted = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), dateFormat);
      blackoutDateMap[formatted] = j;
    } else {
      blackoutDateMap[d] = j;
    }
  }
  var volunteerRowMap = {};
  for (var i = 1; i < blackoutData.length; i++) {
    var volName = blackoutData[i][0];
    volunteerRowMap[volName] = i;
  }

  /**
   * Set per-cell data validation on the Schedule sheet so that:
   * - Each dropdown only shows volunteers who are eligible for that role
   * - AND are NOT blacked-out on that specific Sunday.
   */
  sundays.forEach(function (dateObj, rIndex) {
    var formattedDate = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), dateFormat);

    rolesHeader.forEach(function (role, cIndex) {
      var baseList = roleVolunteers[role] || [];

      // Filter out volunteers who have blackout === TRUE on this date.
      var filteredList = baseList.filter(function (volName) {
        // If we have no blackout row or no column for this date, treat as available.
        if (!volunteerRowMap.hasOwnProperty(volName) || !blackoutDateMap.hasOwnProperty(formattedDate)) {
          return true;
        }

        var rowIdx = volunteerRowMap[volName];      // index into blackoutData rows
        var colIdx = blackoutDateMap[formattedDate]; // index into blackoutData columns
        var cellVal = blackoutData[rowIdx][colIdx];

        // If the cell is TRUE, it means the volunteer is blacked out -> exclude.
        return cellVal !== true;
      });

      var cell = scheduleSheet.getRange(rIndex + 2, cIndex + 2); // +2 because row 2/col 2 are first data cells

      if (filteredList.length > 0) {
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(filteredList, true)
          .build();
        cell.setDataValidation(rule);
      } else {
        // No available volunteers for this role on this date: clear validation.
        cell.clearDataValidations();
      }
    });
  });

  // Set up a persistent round-robin pointer for each role.
  var lastAssignedIndex = {};
  rolesHeader.forEach(function (role) {
    lastAssignedIndex[role] = -1;
  });

  var floatingRoles = getFloatingRoles();

  // Map each volunteer to their spouse (if any), from Couples sheet.
  var couplesMap = getCouplesMap();

  // Track who served on the previous Sunday (any role).
  var servedLastSunday = {};

  // For each Sunday (each row in Schedule starting at row 2) and each role,
  // assign a volunteer using round-robin that respects:
  //   - blackout dates
  //   - one non-floating role per person per Sunday
  //   - couples cannot serve on the same Sunday
  //   - no one serves two consecutive Sundays
  for (var r = 0; r < sundays.length; r++) {
    var currentSunday = sundays[r];
    var currentSundayFormatted = Utilities.formatDate(currentSunday, ss.getSpreadsheetTimeZone(), dateFormat);

    // Track volunteers already assigned on THIS date (any role)
    var assignedForDate = [];

    rolesHeader.forEach(function (role, i) {
      var volunteers = roleVolunteers[role];
      var assigned = "";
      var isFloating = floatingRoles.indexOf(role) !== -1;

      if (volunteers.length > 0) {
        // Start from the volunteer following the last assigned one for this role.
        var startIndex = (lastAssignedIndex[role] + 1) % volunteers.length;
        var candidate = null;

        for (var k = 0; k < volunteers.length; k++) {
          var index = (startIndex + k) % volunteers.length;
          var volName = volunteers[index];

          // 1) Skip if volunteer is already assigned a non-floating role on this date.
          if (!isFloating && assignedForDate.indexOf(volName) !== -1) {
            continue;
          }

          // 2) Skip if this volunteer served last Sunday (no back-to-back Sundays).
          if (servedLastSunday[volName]) {
            continue;
          }

          // 3) Skip if volunteer's spouse is already serving on this date.
          var spouse = couplesMap[volName];
          if (spouse && assignedForDate.indexOf(spouse) !== -1) {
            continue;
          }

          // 4) Check if volunteer has a blackout on this date.
          var isBlackout = false;
          if (volunteerRowMap.hasOwnProperty(volName) && blackoutDateMap.hasOwnProperty(currentSundayFormatted)) {
            var bdValue = blackoutData[volunteerRowMap[volName]][blackoutDateMap[currentSundayFormatted]];
            if (bdValue === true) {
              isBlackout = true;
            }
          }

          // 5) If passes all checks, pick this volunteer.
          if (!isBlackout) {
            candidate = volName;
            lastAssignedIndex[role] = index; // update the pointer for this role
            break;
          }
        }

        if (candidate) {
          assigned = candidate;
          assignedForDate.push(candidate);
        }
      }

      // Write assignment (or blank if no valid candidate)
      scheduleSheet.getRange(r + 2, i + 2).setValue(assigned);
    });

    // After finishing this Sunday, update servedLastSunday for the next iteration.
    servedLastSunday = {};
    assignedForDate.forEach(function (name) {
      servedLastSunday[name] = true;
    });
  }

  SpreadsheetApp.getUi().alert("Schedule auto-populated successfully.");
}

/**
 * An installable onEdit trigger for logging changes in the Schedule sheet.
 * When a manual edit occurs, logs a message like:
 * "Position [role] is changed from [oldValue] to [newValue] for the date [formattedDate]".
 *
 * Only manual changes trigger this event. Programmatic changes (like autoPopulateSchedule) are ignored.
 */
function handleScheduleEdit(e) {
  // Ensure the event object is present.
  if (!e) return;

  var sheet = e.range.getSheet();

  // Only proceed if the edited sheet is "Schedule".
  if (sheet.getName() !== "Schedule") return;

  // Ignore edits in the header row or the first column (date column).
  if (e.range.getRow() < 2 || e.range.getColumn() < 2) return;

  // Get the role name from the header (row 1) at the edited column.
  var role = sheet.getRange(1, e.range.getColumn()).getValue();

  // Retrieve the date from column A in the same row.
  var dateCell = sheet.getRange(e.range.getRow(), 1).getValue();
  if (!(dateCell instanceof Date)) return; // if no valid date, skip.

  // Format the date in Pacific Time (yyyy/MM/dd).
  var formattedDate = Utilities.formatDate(new Date(dateCell), "America/Los_Angeles", "yyyy/MM/dd");

  // Retrieve the old and new values. (e.oldValue is only available with an installable trigger.)
  var oldValue = e.oldValue || "";
  var newValue = e.value || "";

  // If there is no change, exit.
  if (oldValue === newValue) return;

  // Build the log message.
  var description = "**" + role + "** is changed from (" + oldValue + ") to (" + newValue + ") for " + formattedDate;

  // Log the action using the logAction() function.
  logAction(description);
}

/**
 * Copies the current Schedule into Schedule History,
 * overwriting any existing rows for the same quarter.
 */
function copyScheduleToHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var scheduleSheet = ss.getSheetByName("Schedule");
  var historySheet = ss.getSheetByName("Schedule History");
  var scheduleData = scheduleSheet.getDataRange().getValues();

  // 1) Create history sheet if needed, and set header row
  if (!historySheet) {
    historySheet = ss.insertSheet("Schedule History");
    historySheet
      .getRange(1, 1, 1, scheduleData[0].length)
      .setValues([scheduleData[0]]);
  }

  // 2) Determine the next-quarter range
  var qr = getNextQuarterRange();
  var startDate = qr.startDate;
  var endDate = qr.endDate;

  // 3) Remove any existing history rows for that quarter
  var historyData = historySheet.getDataRange().getValues();
  for (var i = historyData.length - 1; i >= 0; i--) {
    var rowDate = historyData[i][0];
    if (rowDate instanceof Date &&
      rowDate >= startDate &&
      rowDate <= endDate) {
      historySheet.deleteRow(i + 1);
    }
  }
  // Remove the left over header row
  historyData = historySheet.getDataRange().getValues();
  if (historyData.length > 0) {
    if (historyData[historyData.length - 1][0] === 'Date')
      historySheet.deleteRow(historyData.length);
  }

  // 4) Append all Schedule rows (skip header at index 0)
  for (var r = 0; r < scheduleData.length; r++) {
    historySheet.appendRow(scheduleData[r]);
  }

  // 5) Notify
  SpreadsheetApp.getUi().alert(
    'Schedule History updated for ' +
    Utilities.formatDate(startDate, tz, 'yyyy/MM/dd') +
    ' – ' +
    Utilities.formatDate(endDate, tz, 'yyyy/MM/dd')
  );
}

/**
 * Highlights three types of conflicts on the Schedule sheet:
 *
 *  1) Same person assigned more than once on the same Sunday (same row, different roles)
 *     -> Light red (#FFCCCC)
 *
 *  2) Same person assigned on two consecutive Sundays (adjacent rows, any roles)
 *     -> Light yellow (#FFF2CC)
 *
 *  3) Husband and wife serving on the same Sunday (from Couples sheet)
 *     -> Light blue (#CCE5FF)
 *
 * Colors are layered with simple priority:
 *   - Same-day duplicate (red) is applied first
 *   - Consecutive-week conflict (yellow) can override red
 *   - Couple conflict (blue) can override both (highest priority)
 */
function highlightConflicts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Schedule");
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return;

  // Data range: from row 2 (first Sunday) and column 2 (first role)
  var numRows = lastRow - 1;
  var numCols = lastCol - 1;
  var range = sheet.getRange(2, 2, numRows, numCols);
  var values = range.getValues();

  // Clear previous backgrounds
  range.setBackground(null);

  // Conflict flags: same-day duplicates, consecutive weeks, couples same day
  var sameDayDup = [];
  var consecWeekDup = [];
  var coupleConflict = [];

  for (var r = 0; r < numRows; r++) {
    sameDayDup[r] = [];
    consecWeekDup[r] = [];
    coupleConflict[r] = [];
    for (var c = 0; c < numCols; c++) {
      sameDayDup[r][c] = false;
      consecWeekDup[r][c] = false;
      coupleConflict[r][c] = false;
    }
  }

  // --- 1) Same-day duplicates (existing behavior, but now via arrays) ---
  for (var r = 0; r < numRows; r++) {
    var counts = {};
    // Count occurrences per name in this row
    for (var c = 0; c < numCols; c++) {
      var name = values[r][c];
      if (name && name !== "NA") {
        counts[name] = (counts[name] || 0) + 1;
      }
    }
    // Mark cells where the name appears more than once
    for (var c = 0; c < numCols; c++) {
      var name = values[r][c];
      if (name && counts[name] > 1) {
        sameDayDup[r][c] = true;
      }
    }
  }

  // --- 2) Duplicates on two consecutive Sundays (adjacent rows) ---
  // For each pair of consecutive rows r and r+1, if a name appears in both,
  // mark all occurrences of that name in both rows.
  for (var r = 0; r < numRows - 1; r++) {
    var rowNow = values[r];
    var rowNext = values[r + 1];

    var namesNow = {};
    var namesNext = {};

    // Collect names in current row
    for (var c = 0; c < numCols; c++) {
      var name = rowNow[c];
      if (name && name !== "NA") namesNow[name] = true;
    }

    // Collect names in next row
    for (var c = 0; c < numCols; c++) {
      var name = rowNext[c];
      if (name && name !== "NA") namesNext[name] = true;
    }

    // Intersection: names serving on consecutive Sundays
    for (var name in namesNow) {
      if (namesNext[name]) {
        // Mark all occurrences in row r
        for (var c = 0; c < numCols; c++) {
          if (values[r][c] === name) {
            consecWeekDup[r][c] = true;
          }
        }
        // Mark all occurrences in row r+1
        for (var c = 0; c < numCols; c++) {
          if (values[r + 1][c] === name) {
            consecWeekDup[r + 1][c] = true;
          }
        }
      }
    }
  }

  // --- 3) Husband & wife on the same Sunday ---
  // Use Couples sheet via getCouplesMap()
  var couplesMap = getCouplesMap();  // { "HusbandName": "WifeName", "WifeName": "HusbandName", ... }

  for (var r = 0; r < numRows; r++) {
    var rowValues = values[r];
    var rowNames = {};

    // Collect who is serving this Sunday
    for (var c = 0; c < numCols; c++) {
      var name = rowValues[c];
      if (name && name !== "NA") {
        rowNames[name] = true;
      }
    }

    // For each cell, if this name has a spouse also in this row, mark as couple conflict
    for (var c = 0; c < numCols; c++) {
      var name = rowValues[c];
      if (!name) continue;

      var spouse = couplesMap[name];
      if (spouse && rowNames[spouse]) {
        coupleConflict[r][c] = true;
      }
    }
  }

  // --- Apply background colors based on conflicts ---
  // We'll build a 2D array of colors to set in one go.
  var colors = [];
  for (var r = 0; r < numRows; r++) {
    colors[r] = [];
    for (var c = 0; c < numCols; c++) {
      var color = null;

      if (sameDayDup[r][c]) {
        color = "#FFCCCC"; // light red: same-day multiple roles
      }
      if (consecWeekDup[r][c]) {
        color = "#FFF2CC"; // light yellow: consecutive Sundays
      }
      if (coupleConflict[r][c]) {
        color = "#CCE5FF"; // light blue: couple serving same day (highest priority)
      }

      colors[r][c] = color;
    }
  }

  range.setBackgrounds(colors);
  SpreadsheetApp.flush();

  Utilities.sleep(10000);

  // Restore to no background
  range.setBackground(null);
  SpreadsheetApp.flush();
}

/**
 * Standalone one-time function to add "NA" option to all existing dropdowns on the Schedule sheet.
 */
function addNaToDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Schedule");
  if (!sheet) {
    if (SpreadsheetApp.getUi()) {
      SpreadsheetApp.getUi().alert("Schedule sheet not found.");
    } else {
      Logger.log("Schedule sheet not found.");
    }
    return;
  }

  var range = sheet.getDataRange();
  var validations = range.getDataValidations();
  var updatedValidations = [];
  var hasUpdates = false;

  for (var i = 0; i < validations.length; i++) {
    var rowRules = [];
    for (var j = 0; j < validations[i].length; j++) {
      var rule = validations[i][j];

      if (rule != null && rule.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        var args = rule.getCriteriaValues();
        var values = args[0]; // The list of values

        // Add "NA" if not present
        if (values.indexOf("NA") === -1) {
          values.push("NA");
          var newRule = SpreadsheetApp.newDataValidation()
            .requireValueInList(values, true)
            .build();
          rowRules.push(newRule);
          hasUpdates = true;
        } else {
          rowRules.push(rule);
        }
      } else {
        rowRules.push(rule);
      }
    }
    updatedValidations.push(rowRules);
  }

  if (hasUpdates) {
    range.setDataValidations(updatedValidations);
    if (SpreadsheetApp.getUi()) {
      SpreadsheetApp.getUi().alert("Added 'NA' option to dropdowns.");
    }
  } else {
    if (SpreadsheetApp.getUi()) {
      SpreadsheetApp.getUi().alert("No dropdowns needed updating.");
    }
  }
}

/**
 * Menu handler: Highlights the selected person and their spouse across the Schedule sheet.
 * Called from the custom menu "Highlight One Person".
 */
function highlightOnePerson() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Check if on Schedule sheet
  if (sheet.getName() !== "Schedule") {
    ui.alert("Please select a cell on the Schedule sheet.");
    return;
  }

  var selection = sheet.getActiveRange();
  if (!selection) {
    ui.alert("Please select a data cell on the Schedule sheet (not the header row or Date column).");
    return;
  }

  // Check if selection is in data area (row >= 2, col >= 2)
  var row = selection.getRow();
  var col = selection.getColumn();

  if (row < 2 || col < 2) {
    ui.alert("Please select a data cell (not the header row or Date column).");
    return;
  }

  var val = selection.getValue();

  // Check if cell is empty or NA
  if (!val || (typeof val === 'string' && val.trim() === "") || val === "NA") {
    ui.alert("Please select a cell with a person's name (not blank or 'NA').");
    return;
  }

  var person = String(val).trim();

  // Retrieve spouse if any
  var couplesMap = getCouplesMap();
  var spouse = couplesMap[person];

  // Get data range (skip header row and date column)
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 2) {
    ui.alert("No data found on the Schedule sheet.");
    return;
  }

  var dataRange = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
  var values = dataRange.getValues();
  var backgrounds = dataRange.getBackgrounds();

  // Collect cells to highlight with their original backgrounds
  var cellsToHighlight = [];

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var cellVal = values[r][c];

      if (typeof cellVal === 'string' && cellVal !== "") {
        if (cellVal === person || (spouse && cellVal === spouse)) {
          cellsToHighlight.push({
            row: r,
            col: c,
            originalBg: backgrounds[r][c]
          });
        }
      }
    }
  }

  if (cellsToHighlight.length === 0) {
    ui.alert("No matching cells found for '" + person + "'.");
    return;
  }

  // Apply Pink Highlight
  var HIGHLIGHT = "#FFCCCC"; // Light Pink

  cellsToHighlight.forEach(function (cell) {
    backgrounds[cell.row][cell.col] = HIGHLIGHT;
  });

  dataRange.setBackgrounds(backgrounds);
  SpreadsheetApp.flush();

  Utilities.sleep(5000);

  // Restore original backgrounds
  cellsToHighlight.forEach(function (cell) {
    backgrounds[cell.row][cell.col] = cell.originalBg;
  });

  dataRange.setBackgrounds(backgrounds);
  SpreadsheetApp.flush();
}
