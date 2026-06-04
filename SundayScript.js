/**
 * Creates a blank schedule for the next quarter.
 * - Preserves existing schedule data and appends new Sundays.
 * - Writes dates into column A.
 * - Populates dropdown options for each cell based on role eligibility and blackout dates.
 * - After this step, admin can manually pre-schedule some roles before running autoPopulateSchedule().
 */
function createBlankSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("Schedule");
  var rolesSheet = ss.getSheetByName("Roles");
  var blackoutSheet = ss.getSheetByName("Blackout Dates");
  var dateFormat = "MM/dd/yyyy";
  var tz = ss.getSpreadsheetTimeZone();

  // Get all Sundays for the next quarter.
  var sundays = getSundaysForNextQuarter();

  // Get role headers from the Roles sheet.
  var lastCol = rolesSheet.getLastColumn();
  var rolesHeader = rolesSheet.getRange(1, 2, 1, lastCol - 2).getValues()[0];

  // Build mapping of roles to qualified volunteers.
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

  // --- Determine where to start appending ---
  var existingLastRow = scheduleSheet.getLastRow();
  var startRowForNewData = 2;

  if (existingLastRow >= 2) {
    startRowForNewData = existingLastRow + 1;
  } else {
    var scheduleHeader = ["Date"].concat(rolesHeader);
    var headerRange = scheduleSheet.getRange(1, 1, 1, scheduleHeader.length);
    headerRange.setValues([scheduleHeader]);
    headerRange.setBackground("#CCCCCC");
  }

  // Clear rows we're about to write
  var maxCols = scheduleSheet.getMaxColumns();
  if (maxCols < rolesHeader.length + 1) {
    maxCols = rolesHeader.length + 1;
  }
  var newRowsRange = scheduleSheet.getRange(startRowForNewData, 1, sundays.length, maxCols);
  newRowsRange.clearContent();
  newRowsRange.clearFormat();
  newRowsRange.clearDataValidations();

  // Write Sunday dates into column A
  var dateValues = sundays.map(function (date) { return [date]; });
  scheduleSheet.getRange(startRowForNewData, 1, sundays.length, 1).setValues(dateValues);
  scheduleSheet.getRange(startRowForNewData, 1, sundays.length, 1).setNumberFormat(dateFormat);
  scheduleSheet.getRange(startRowForNewData, 1, sundays.length, 1).setBackground("#DDDDDD");

  // Load blackout data.
  var blackoutData = blackoutSheet.getDataRange().getValues();
  var blackoutHeader = blackoutData[0];
  var blackoutDateMap = {};
  for (var j = 1; j < blackoutHeader.length; j++) {
    var d = blackoutHeader[j];
    if (d instanceof Date) {
      blackoutDateMap[Utilities.formatDate(d, tz, dateFormat)] = j;
    } else {
      blackoutDateMap[d] = j;
    }
  }
  var volunteerRowMap = {};
  for (var i = 1; i < blackoutData.length; i++) {
    volunteerRowMap[blackoutData[i][0]] = i;
  }

  // Get Schedule sheet headers and build role -> column index map
  var scheduleHeaders = scheduleSheet.getRange(1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
  var roleToColumnIndex = {};
  for (var h = 1; h < scheduleHeaders.length; h++) {
    var headerRole = (scheduleHeaders[h] || "").toString().trim();
    if (headerRole) {
      roleToColumnIndex[headerRole] = h + 1;
    }
  }

  // Set per-cell data validation for each new Sunday
  sundays.forEach(function (dateObj, rIndex) {
    var formattedDate = Utilities.formatDate(dateObj, tz, dateFormat);
    rolesHeader.forEach(function (role) {
      var baseList = roleVolunteers[role] || [];
      var colIndex = roleToColumnIndex[role];
      if (!colIndex) return;

      var filteredList = baseList.filter(function (volName) {
        if (!volunteerRowMap.hasOwnProperty(volName) || !blackoutDateMap.hasOwnProperty(formattedDate)) return true;
        return blackoutData[volunteerRowMap[volName]][blackoutDateMap[formattedDate]] !== true;
      });

      var cell = scheduleSheet.getRange(startRowForNewData + rIndex, colIndex);
      var dropdownList = filteredList.slice();
      if (dropdownList.indexOf("NA") === -1) dropdownList.push("NA");
      cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(dropdownList, true).build());
    });
  });

  SpreadsheetApp.getUi().alert("Blank schedule created. Added " + sundays.length + " new Sundays starting at row " + startRowForNewData + ". You may now manually pre-schedule before running Auto Populate.");
}

/**
 * Auto-populates the Schedule sheet with volunteer assignments.
 * Should be run AFTER createBlankSchedule() and any manual pre-scheduling.
 * - Reads existing rows to determine the new quarter's date range.
 * - Skips cells that already have a manual assignment (allows admin pre-scheduling).
 * - Pre-scheduled volunteers are excluded from all auto-scheduling.
 * - On Combined Dates, assigns "大堂 Combine" to "Lion Teacher".
 * - Schedules "double week roles" first: same volunteer serves 2 consecutive Sundays.
 * - Handles cross-quarter boundary for double-week roles (carry-over from last quarter).
 * - Ensures no volunteer is assigned more than one non-floating role on the same day.
 * - Ensures no volunteer (or their spouse) serves on the same day.
 * - After a double-week pair, the volunteer is blocked from ALL roles the following week.
 */
function autoPopulateSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("Schedule");
  var rolesSheet = ss.getSheetByName("Roles");
  var blackoutSheet = ss.getSheetByName("Blackout Dates");
  var dateFormat = "MM/dd/yyyy";
  var tz = ss.getSpreadsheetTimeZone();

  // Get role headers from the Roles sheet.
  var lastCol = rolesSheet.getLastColumn();
  var rolesHeader = rolesSheet.getRange(1, 2, 1, lastCol - 2).getValues()[0];

  // Build mapping of roles to qualified volunteers.
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

  // Remove placeholder names (e.g. "大堂 Combine") from all volunteer lists.
  // These are not real volunteers and should never be auto-assigned.
  var COMBINE_NAME = "大堂 Combine";
  rolesHeader.forEach(function (role) {
    roleVolunteers[role] = roleVolunteers[role].filter(function (name) {
      return name.indexOf("大堂") === -1;
    });
  });

  // --- Determine the new quarter's Sundays from the Schedule sheet ---
  // Read all dates from the Schedule sheet and find Sundays for the next quarter.
  var sundays = getSundaysForNextQuarter();
  var sundayStrSet = {};
  sundays.forEach(function (d) {
    sundayStrSet[Utilities.formatDate(d, tz, dateFormat)] = true;
  });

  // Find the first row of the new quarter on the Schedule sheet
  var allScheduleData = scheduleSheet.getDataRange().getValues();
  var startRowForNewData = -1;
  for (var i = 1; i < allScheduleData.length; i++) {
    var cellDate = allScheduleData[i][0];
    if (cellDate instanceof Date) {
      var formatted = Utilities.formatDate(cellDate, tz, dateFormat);
      if (sundayStrSet[formatted]) {
        startRowForNewData = i + 1; // 1-indexed sheet row
        break;
      }
    }
  }

  if (startRowForNewData === -1) {
    SpreadsheetApp.getUi().alert("Could not find next quarter's dates on the Schedule sheet. Please run 'Create Blank Schedule' first.");
    return;
  }

  // --- Read last 2 rows before the new quarter for carry-over ---
  var scheduleHeaders = scheduleSheet.getRange(1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
  var lastSundayAssignments = {};
  var secondLastSundayAssignments = {};

  if (startRowForNewData > 2) {
    var lastRowData = scheduleSheet.getRange(startRowForNewData - 1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
    for (var c = 1; c < scheduleHeaders.length; c++) {
      var roleName = (scheduleHeaders[c] || "").toString().trim();
      var assignedVol = (lastRowData[c] || "").toString().trim();
      if (roleName && assignedVol && assignedVol !== "NA") {
        lastSundayAssignments[roleName] = assignedVol;
      }
    }
  }

  if (startRowForNewData > 3) {
    var secondLastRowData = scheduleSheet.getRange(startRowForNewData - 2, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
    for (var c = 1; c < scheduleHeaders.length; c++) {
      var roleName = (scheduleHeaders[c] || "").toString().trim();
      var assignedVol = (secondLastRowData[c] || "").toString().trim();
      if (roleName && assignedVol && assignedVol !== "NA") {
        secondLastSundayAssignments[roleName] = assignedVol;
      }
    }
  }

  // Load blackout data.
  var blackoutData = blackoutSheet.getDataRange().getValues();
  var blackoutHeader = blackoutData[0];
  var blackoutDateMap = {};
  for (var j = 1; j < blackoutHeader.length; j++) {
    var d = blackoutHeader[j];
    if (d instanceof Date) {
      blackoutDateMap[Utilities.formatDate(d, tz, dateFormat)] = j;
    } else {
      blackoutDateMap[d] = j;
    }
  }
  var volunteerRowMap = {};
  for (var i = 1; i < blackoutData.length; i++) {
    volunteerRowMap[blackoutData[i][0]] = i;
  }

  // Build role -> column index map
  var roleToColumnIndex = {};
  for (var h = 1; h < scheduleHeaders.length; h++) {
    var headerRole = (scheduleHeaders[h] || "").toString().trim();
    if (headerRole) {
      roleToColumnIndex[headerRole] = h + 1;
    }
  }

  // --- Load config ---
  var floatingRoles = getFloatingRoles();
  var doubleWeekRoles = getDoubleWeekRoles();
  var couplesMap = getCouplesMap();

  var doubleWeekSet = {};
  doubleWeekRoles.forEach(function (r) { doubleWeekSet[r] = true; });

  // --- Initialize round-robin pointers ---
  var lastAssignedIndex = {};
  rolesHeader.forEach(function (role) {
    var volunteers = roleVolunteers[role] || [];
    var lastVolunteer = lastSundayAssignments[role];
    if (lastVolunteer && volunteers.length > 0) {
      var idx = volunteers.indexOf(lastVolunteer);
      lastAssignedIndex[role] = (idx !== -1) ? idx : -1;
    } else {
      lastAssignedIndex[role] = -1;
    }
  });

  // --- Build tracking arrays ---
  var assignments = [];
  var servedOnWeek = [];
  for (var s = 0; s < sundays.length; s++) {
    assignments[s] = {};
    servedOnWeek[s] = {};
  }

  // Previous quarter served data
  var prevWeekServed = [{}, {}];
  for (var role in lastSundayAssignments) {
    prevWeekServed[1][lastSundayAssignments[role]] = true;
  }
  for (var role in secondLastSundayAssignments) {
    prevWeekServed[0][secondLastSundayAssignments[role]] = true;
  }

  // --- Phase 1: Read manual pre-assignments ---
  var newDataRange = scheduleSheet.getRange(startRowForNewData, 1, sundays.length, scheduleSheet.getLastColumn());
  var newDataValues = newDataRange.getValues();
  var preScheduledVolunteers = {};

  for (var r = 0; r < sundays.length; r++) {
    for (var h = 1; h < scheduleHeaders.length; h++) {
      var role = (scheduleHeaders[h] || "").toString().trim();
      if (!role) continue;
      var cellVal = (newDataValues[r][h] || "").toString().trim();
      if (cellVal && cellVal !== "NA") {
        assignments[r][role] = cellVal;
        servedOnWeek[r][cellVal] = true;
        preScheduledVolunteers[cellVal] = true;
      }
    }
  }

  // --- Phase 1.5: Assign "大堂 Combine" to "Lion Teacher" on Combined Dates ---
  var configSheet = ss.getSheetByName("Config");
  var combinedDatesSet = {};
  if (configSheet) {
    var cfgLastRow = configSheet.getLastRow();
    if (cfgLastRow >= 2) {
      var combinedValues = configSheet.getRange(2, 5, cfgLastRow - 1, 1).getValues();
      for (var i = 0; i < combinedValues.length; i++) {
        var cv = combinedValues[i][0];
        if (cv instanceof Date) {
          combinedDatesSet[Utilities.formatDate(cv, tz, dateFormat)] = true;
        }
      }
    }
  }

  var LION_TEACHER_ROLE = "Lion Teacher";
  if (roleToColumnIndex[LION_TEACHER_ROLE]) {
    for (var r = 0; r < sundays.length; r++) {
      var sundayStr = Utilities.formatDate(sundays[r], tz, dateFormat);
      if (combinedDatesSet[sundayStr] && !assignments[r][LION_TEACHER_ROLE]) {
        assignments[r][LION_TEACHER_ROLE] = COMBINE_NAME;
        // Note: do NOT add COMBINE_NAME to servedOnWeek — it's a placeholder, not a real volunteer
      }
    }
  }
  preScheduledVolunteers[COMBINE_NAME] = true;

  // --- Helper functions ---
  function isBlackout(volName, sundayIdx) {
    var formattedDate = Utilities.formatDate(sundays[sundayIdx], tz, dateFormat);
    if (volunteerRowMap.hasOwnProperty(volName) && blackoutDateMap.hasOwnProperty(formattedDate)) {
      return blackoutData[volunteerRowMap[volName]][blackoutDateMap[formattedDate]] === true;
    }
    return false;
  }

  function didServe(volName, weekIdx) {
    if (weekIdx >= 0) {
      return !!servedOnWeek[weekIdx][volName];
    } else if (weekIdx === -1) {
      return !!prevWeekServed[1][volName];
    } else if (weekIdx === -2) {
      return !!prevWeekServed[0][volName];
    }
    return false;
  }

  function wouldCauseThreeConsecutive(volName, weekIdx) {
    if (didServe(volName, weekIdx - 1) && didServe(volName, weekIdx - 2)) return true;
    if (didServe(volName, weekIdx - 1) && weekIdx + 1 < sundays.length && didServe(volName, weekIdx + 1)) return true;
    if (weekIdx + 1 < sundays.length && didServe(volName, weekIdx + 1) &&
        weekIdx + 2 < sundays.length && didServe(volName, weekIdx + 2)) return true;
    return false;
  }

  function pickVolunteer(role, sundayIdx, isDoubleWeek) {
    var volunteers = roleVolunteers[role] || [];
    if (volunteers.length === 0) return "NA";

    var isFloating = floatingRoles.indexOf(role) !== -1;
    var startIndex = (lastAssignedIndex[role] + 1) % volunteers.length;

    for (var k = 0; k < volunteers.length; k++) {
      var index = (startIndex + k) % volunteers.length;
      var volName = volunteers[index];

      if (preScheduledVolunteers[volName]) continue;
      if (!isFloating && servedOnWeek[sundayIdx][volName]) continue;
      var spouse = couplesMap[volName];
      if (spouse && servedOnWeek[sundayIdx][spouse]) continue;
      if (isBlackout(volName, sundayIdx)) continue;
      // Skip if volunteer served previous week (back-to-back).
      // For double-week: a 2-week pair starting here would cause 3 consecutive weeks.
      // For regular: no back-to-back Sundays.
      if (didServe(volName, sundayIdx - 1)) continue;
      if (wouldCauseThreeConsecutive(volName, sundayIdx)) continue;

      lastAssignedIndex[role] = index;
      return volName;
    }
    return "NA";
  }

  function recordAssignment(sundayIdx, role, volName) {
    assignments[sundayIdx][role] = volName;
    if (volName && volName !== "NA" && !isPlaceholder(volName)) {
      servedOnWeek[sundayIdx][volName] = true;
    }
  }

  // --- Helper: check if a value is a placeholder (like 大堂 Combine) rather than a real volunteer ---
  function isPlaceholder(val) {
    if (!val) return false;
    return val === COMBINE_NAME || val.indexOf("大堂") !== -1;
  }

  // --- Phase 2: Schedule double-week roles first ---
  // Double-week roles assign the same volunteer for 2 consecutive Sundays.
  // Combined dates ("大堂 Combine") are treated as gaps — not real assignments.
  // If volunteer A serves week N, and week N+1 is a combined date, A carries over to week N+2.
  doubleWeekRoles.forEach(function (role) {
    if (!roleToColumnIndex[role]) return;
    if (!roleVolunteers[role] || roleVolunteers[role].length === 0) return;

    // Determine carry-over from previous quarter
    var carryOver = null; // volunteer who needs a second week
    var lastVol = lastSundayAssignments[role];
    var secondLastVol = secondLastSundayAssignments[role];

    if (lastVol && lastVol !== "NA" && !isPlaceholder(lastVol)) {
      // If last week's volunteer differs from second-to-last, they only served 1 week — carry over
      if (secondLastVol !== lastVol) {
        carryOver = lastVol;
      }
    }

    // Iterate week by week, tracking how many weeks the current volunteer has served
    var currentVol = carryOver;  // volunteer currently being paired
    var weeksServed = carryOver ? 1 : 0;  // they already served 1 week in previous quarter

    for (var w = 0; w < sundays.length; w++) {
      // If this week is already assigned
      if (assignments[w][role]) {
        var existing = assignments[w][role];
        if (isPlaceholder(existing)) {
          // Combined date — treat as gap, don't reset the current volunteer's pairing
          continue;
        }
        // A real manual assignment — reset pairing to this person
        currentVol = existing;
        weeksServed = 1;
        continue;
      }

      // Need to auto-assign this week
      if (currentVol && weeksServed === 1 && !isPlaceholder(currentVol)) {
        // Try to assign the same volunteer for their second week
        if (!isBlackout(currentVol, w) && !wouldCauseThreeConsecutive(currentVol, w)) {
          var isFloating = floatingRoles.indexOf(role) !== -1;
          var spouseW = couplesMap[currentVol];
          var spouseConflict = spouseW && servedOnWeek[w][spouseW];
          var alreadyAssigned = !isFloating && servedOnWeek[w][currentVol];
          if (!spouseConflict && !alreadyAssigned) {
            recordAssignment(w, role, currentVol);
            weeksServed = 2;
            continue;
          }
        }
        // Can't assign same volunteer — pick a new one for a fresh pair
      }

      // Pick a new volunteer for a fresh 2-week pair
      var vol = pickVolunteer(role, w, true);
      recordAssignment(w, role, vol);
      currentVol = (vol !== "NA") ? vol : null;
      weeksServed = (vol !== "NA") ? 1 : 0;
    }
  });

  // --- Phase 3: Schedule remaining (non-double-week) roles ---
  for (var r = 0; r < sundays.length; r++) {
    rolesHeader.forEach(function (role) {
      if (doubleWeekSet[role]) return;
      var colIndex = roleToColumnIndex[role];
      if (!colIndex) return;
      if (assignments[r][role]) return;

      if (role.indexOf("Parent Helper") !== -1) {
        recordAssignment(r, role, "NA");
        return;
      }

      var vol = pickVolunteer(role, r, false);
      recordAssignment(r, role, vol);
    });
  }

  // --- Write all assignments to the sheet ---
  for (var r = 0; r < sundays.length; r++) {
    for (var role in assignments[r]) {
      var colIndex = roleToColumnIndex[role];
      if (colIndex && assignments[r][role]) {
        scheduleSheet.getRange(startRowForNewData + r, colIndex).setValue(assignments[r][role]);
      }
    }
  }

  SpreadsheetApp.getUi().alert("Schedule auto-populated successfully for " + sundays.length + " Sundays starting at row " + startRowForNewData + ".");
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

  // If the edited row's date is the upcoming Sunday, highlight the cell light green
  // (same as the "confirmed" color from the web app).
  try {
    var upcomingDate = findUpcomingSundayDate_(sheet);
    if (upcomingDate) {
      var rowDate = new Date(dateCell.getTime());
      rowDate.setHours(0, 0, 0, 0);
      upcomingDate.setHours(0, 0, 0, 0);
      if (rowDate.getTime() === upcomingDate.getTime() && newValue && newValue !== "NA") {
        e.range.setBackground(CONFIRM_COLOR);
      }
    }
  } catch (highlightErr) {
    Logger.log("Error in upcoming Sunday highlight: " + highlightErr.message);
  }

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
 *  2) Same person serving 3 consecutive Sundays
 *     -> Light yellow (#FFF2CC)
 *
 *  3) Husband and wife serving on the same Sunday (from Couples sheet)
 *     -> Light blue (#CCE5FF)
 *
 * Colors are layered with simple priority:
 *   - Same-day duplicate (red) is applied first
 *   - 3-consecutive-week conflict (yellow) can override red
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

  // Save existing backgrounds so we can restore them after highlighting
  var savedBackgrounds = range.getBackgrounds();

  // Conflict flags: same-day duplicates, 3-consecutive-weeks, couples same day
  var sameDayDup = [];
  var threeConsec = [];
  var coupleConflict = [];

  for (var r = 0; r < numRows; r++) {
    sameDayDup[r] = [];
    threeConsec[r] = [];
    coupleConflict[r] = [];
    for (var c = 0; c < numCols; c++) {
      sameDayDup[r][c] = false;
      threeConsec[r][c] = false;
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

  // --- 2) Same person serving 3 consecutive Sundays ---
  // For each triplet of consecutive rows r, r+1, r+2, if a name appears in all three,
  // mark all occurrences of that name in all three rows.
  for (var r = 0; r < numRows - 2; r++) {
    // Collect names in each of the 3 rows
    var namesRow0 = {};
    var namesRow1 = {};
    var namesRow2 = {};

    for (var c = 0; c < numCols; c++) {
      var n = values[r][c];
      if (n && n !== "NA") namesRow0[n] = true;
      n = values[r + 1][c];
      if (n && n !== "NA") namesRow1[n] = true;
      n = values[r + 2][c];
      if (n && n !== "NA") namesRow2[n] = true;
    }

    // Find names present in all three rows
    for (var name in namesRow0) {
      if (namesRow1[name] && namesRow2[name]) {
        // Mark all occurrences in rows r, r+1, r+2
        for (var c = 0; c < numCols; c++) {
          if (values[r][c] === name) threeConsec[r][c] = true;
          if (values[r + 1][c] === name) threeConsec[r + 1][c] = true;
          if (values[r + 2][c] === name) threeConsec[r + 2][c] = true;
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
  // Start from the saved backgrounds and overlay conflict colors.
  var colors = [];
  for (var r = 0; r < numRows; r++) {
    colors[r] = [];
    for (var c = 0; c < numCols; c++) {
      // Keep the existing background by default
      colors[r][c] = savedBackgrounds[r][c];

      if (sameDayDup[r][c]) {
        colors[r][c] = "#FFCCCC"; // light red: same-day multiple roles
      }
      if (threeConsec[r][c]) {
        colors[r][c] = "#FFF2CC"; // light yellow: 3 consecutive Sundays
      }
      if (coupleConflict[r][c]) {
        colors[r][c] = "#CCE5FF"; // light blue: couple serving same day (highest priority)
      }
    }
  }

  range.setBackgrounds(colors);

  // --- Add legend on rows 35, 36, 37 (columns A & B) ---
  var legendRows = [35, 36, 37];
  var legendColors = ["#FFCCCC", "#FFF2CC", "#CCE5FF"];
  var legendLabels = [
    "Same person in multiple roles on the same Sunday",
    "Same person serving 3 consecutive Sundays",
    "Husband and wife serving on the same Sunday"
  ];

  // Save original values and backgrounds for legend cells
  var legendRange = sheet.getRange(legendRows[0], 1, legendRows.length, 2);
  var savedLegendValues = legendRange.getValues();
  var savedLegendBgs = legendRange.getBackgrounds();

  // Write legend
  for (var i = 0; i < legendRows.length; i++) {
    sheet.getRange(legendRows[i], 1).setValue("").setBackground(legendColors[i]);
    sheet.getRange(legendRows[i], 2).setValue(legendLabels[i]).setBackground(null);
  }

  SpreadsheetApp.flush();

  Utilities.sleep(15000);

  // Restore original backgrounds (preserves confirm/decline colors)
  range.setBackgrounds(savedBackgrounds);

  // Restore legend cells
  legendRange.setValues(savedLegendValues);
  legendRange.setBackgrounds(savedLegendBgs);

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

  Utilities.sleep(10000);

  // Restore original backgrounds
  cellsToHighlight.forEach(function (cell) {
    backgrounds[cell.row][cell.col] = cell.originalBg;
  });

  dataRange.setBackgrounds(backgrounds);
  SpreadsheetApp.flush();
}

/**
 * Installable onEdit trigger for the Roles sheet.
 * When a role checkbox is changed:
 *   - If set to TRUE: add the volunteer to the Schedule dropdowns for that role
 *     (excluding dates where the volunteer has a blackout)
 *   - If set to FALSE: remove the volunteer from Schedule dropdowns for that role
 *     (only for future dates). If volunteer is already assigned, set cell to "NA".
 *
 * Logs all changes to the Roles sheet and any Schedule cell changes to "NA".
 *
 * Must be installed as an installable trigger to get e.oldValue.
 */
function handleRolesEdit(e) {
  if (!e) return;

  var range = e.range;
  var sheet = range.getSheet();

  // Only proceed on the "Roles" sheet
  if (sheet.getName() !== "Roles") return;

  var row = range.getRow();
  var col = range.getColumn();

  // Skip header row (row 1) and name column (col 1)
  // The email column is the last column in the header row, so skip that too
  var headerValues = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];

  // Find the actual last column with data in header row
  var lastDataCol = 1;
  for (var c = 0; c < headerValues.length; c++) {
    if (headerValues[c] !== "") {
      lastDataCol = c + 1; // 1-indexed
    }
  }

  // Skip if: header row, name column (col 1), or email column (last data col)
  if (row < 2 || col < 2 || col >= lastDataCol) return;

  var newValue = e.value;

  // For checkboxes, e.value can be boolean true/false or string "TRUE"/"FALSE"
  // Read the actual cell value to be sure
  var cellValue = range.getValue();
  var isChecked = (cellValue === true);

  // We need to determine if this was a change. Since e.oldValue may not always be available,
  // we'll process the action based on the current state:
  // - If checked (true): add to dropdowns
  // - If unchecked (false): remove from dropdowns
  // The function will be idempotent, so running it multiple times is safe.

  // Get volunteer name from column A
  var volunteerName = sheet.getRange(row, 1).getDisplayValue().trim();
  if (!volunteerName) return;

  // Get role name from header row
  var roleName = sheet.getRange(1, col).getDisplayValue().trim();
  if (!roleName) return;

  // Log the Roles sheet change
  var rolesLogMsg = "**" + roleName + "** role for " + volunteerName + " changed to " + (isChecked ? "enabled" : "disabled");
  logAction(rolesLogMsg);

  // Get necessary sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("Schedule");
  var blackoutSheet = ss.getSheetByName("Blackout Dates");

  if (!scheduleSheet) return;

  // Find the column index on Schedule sheet that matches this role
  var scheduleHeaders = scheduleSheet.getRange(1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
  var roleColIndex = -1;
  for (var i = 0; i < scheduleHeaders.length; i++) {
    if (scheduleHeaders[i] === roleName) {
      roleColIndex = i + 1; // 1-indexed
      break;
    }
  }

  if (roleColIndex < 2) return; // Role not found or it's the Date column

  // Load blackout data to check volunteer's blackout dates
  var blackoutDateMap = {};       // formatted date string -> column index
  var volunteerBlackoutRow = -1;  // row index in blackout data for this volunteer
  var blackoutData = [];
  var dateFormat = "MM/dd/yyyy";
  var tz = ss.getSpreadsheetTimeZone();

  if (blackoutSheet) {
    blackoutData = blackoutSheet.getDataRange().getValues();
    var blackoutHeader = blackoutData[0];

    // Build map of date -> column index
    for (var j = 1; j < blackoutHeader.length; j++) {
      var d = blackoutHeader[j];
      if (d instanceof Date) {
        var formatted = Utilities.formatDate(d, tz, dateFormat);
        blackoutDateMap[formatted] = j;
      } else {
        blackoutDateMap[d] = j;
      }
    }

    // Find volunteer's row in blackout sheet
    for (var i = 1; i < blackoutData.length; i++) {
      if (blackoutData[i][0] === volunteerName) {
        volunteerBlackoutRow = i;
        break;
      }
    }
  }

  // Get today's date for comparison (only modify future dates for FALSE case)
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // Iterate through each date row on Schedule sheet and update dropdowns
  var scheduleData = scheduleSheet.getDataRange().getValues();

  for (var r = 1; r < scheduleData.length; r++) { // Skip header row
    var dateCell = scheduleData[r][0];
    if (!(dateCell instanceof Date)) continue;

    var scheduleDateClean = new Date(dateCell.getTime());
    scheduleDateClean.setHours(0, 0, 0, 0);
    var formattedDate = Utilities.formatDate(dateCell, tz, dateFormat);

    // Get current validation rule for this cell
    var cell = scheduleSheet.getRange(r + 1, roleColIndex); // +1 because r is 0-indexed from data
    var validation = cell.getDataValidation();
    var currentValue = scheduleData[r][roleColIndex - 1]; // roleColIndex is 1-indexed

    if (isChecked) {
      // ========== CHECKBOX SET TO TRUE ==========
      // Check if volunteer has blackout on this date
      var isBlackout = false;
      if (volunteerBlackoutRow >= 0 && blackoutDateMap.hasOwnProperty(formattedDate)) {
        var blackoutColIdx = blackoutDateMap[formattedDate];
        if (blackoutData[volunteerBlackoutRow][blackoutColIdx] === true) {
          isBlackout = true;
        }
      }

      // Skip this date if volunteer has blackout
      if (isBlackout) continue;

      if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        var criteriaValues = validation.getCriteriaValues();
        var currentList = criteriaValues[0]; // Array of allowed values

        // Add volunteer if not already in list
        if (currentList.indexOf(volunteerName) === -1) {
          currentList.push(volunteerName);
          currentList.sort(); // Keep list sorted

          var newRule = SpreadsheetApp.newDataValidation()
            .requireValueInList(currentList, true)
            .build();
          cell.setDataValidation(newRule);
        }
      } else {
        // No existing validation - create one with just this volunteer
        var newRule = SpreadsheetApp.newDataValidation()
          .requireValueInList([volunteerName], true)
          .build();
        cell.setDataValidation(newRule);
      }

    } else {
      // ========== CHECKBOX SET TO FALSE ==========
      // Only modify future dates
      if (scheduleDateClean <= today) continue;

      if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        var criteriaValues = validation.getCriteriaValues();
        var currentList = criteriaValues[0]; // Array of allowed values

        // Remove volunteer from list if present
        var idx = currentList.indexOf(volunteerName);
        if (idx !== -1) {
          currentList.splice(idx, 1);

          if (currentList.length > 0) {
            var newRule = SpreadsheetApp.newDataValidation()
              .requireValueInList(currentList, true)
              .build();
            cell.setDataValidation(newRule);
          } else {
            // No volunteers left - clear validation
            cell.clearDataValidations();
          }
        }
      }

      // If volunteer was already assigned to this cell, set to "NA" and log
      if (currentValue === volunteerName) {
        cell.setValue("NA");
        var scheduleLogMsg = "**" + roleName + "** for " + formattedDate + " changed from (" + volunteerName + ") to (NA) due to role being disabled";
        logAction(scheduleLogMsg);
      }
    }
  }
}
