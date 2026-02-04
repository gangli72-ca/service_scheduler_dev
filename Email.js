/**
 * Sends a sample blackout-dates notification email to the volunteer(s)
 * on the currently selected row(s) in the "Roles" sheet.
 *
 * Assumes:
 *   Roles!A = Name
 *   Roles!last column = Email
 *
 * The email includes a hyperlink that points directly to the
 * "Blackout Dates" sheet for this spreadsheet.
 *
 * Supports:
 *   - Single cell selection on a row
 *   - Selection of multiple rows
 *   - Multiple ranges (using Shift/Ctrl/Cmd selection)
 */
function sendBlackoutNotificationEmails() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rolesSheet = ss.getSheetByName("Roles");
    var blackoutSheet = ss.getSheetByName("Blackout Dates");

    if (!rolesSheet || !blackoutSheet) {
        SpreadsheetApp.getUi().alert('Missing "Roles" or "Blackout Dates" sheet.');
        return;
    }

    var activeSheet = ss.getActiveSheet();
    if (activeSheet.getName() !== "Roles") {
        SpreadsheetApp.getUi().alert('Please select one or more rows on the "Roles" sheet first.');
        return;
    }

    // Gather selected ranges (can be multiple ranges)
    var rangeList = ss.getActiveRangeList();
    if (!rangeList) {
        SpreadsheetApp.getUi().alert('Please select one or more volunteer rows on the "Roles" sheet.');
        return;
    }

    // Find the email column by header name "SVCA Email"
    var headerRow = rolesSheet.getDataRange().getValues()[0];
    var emailCol = -1;
    for (var c = 0; c < headerRow.length; c++) {
        if (String(headerRow[c]).trim() === 'SVCA Email') {
            emailCol = c + 1;
            break;
        }
    }
    if(emailCol == -1) {
        SpreadsheetApp.getUi().alert('Cannot find the SVCA Email column.');
        return;
    }

    var ranges = rangeList.getRanges();
    var recipients = [];  // {name, email, row}

    ranges.forEach(function (range) {
        var startRow = range.getRow();
        var endRow = range.getLastRow();

        for (var r = startRow; r <= endRow; r++) {
            // Skip header row
            if (r < 2) continue;

            var name = rolesSheet.getRange(r, 1).getDisplayValue().trim();        // Col A
            var email = rolesSheet.getRange(r, emailCol).getDisplayValue().trim();  

            if (!name || !email) {
              logAction('Cannot find name or email on Roles sheet for row ' + r);
              continue;
            }
            

            recipients.push({
                name: name,
                email: email,
                row: r
            });
        }
    });

    if (recipients.length === 0) {
        SpreadsheetApp.getUi().alert('No valid name+email pairs found in the selected rows.');
        return;
    }

    // Build a URL that opens directly to the "Blackout Dates" sheet
    var baseUrl = ss.getUrl().split('#')[0];   // strip any existing gid
    var blackoutGid = blackoutSheet.getSheetId();
    var blackoutUrl = baseUrl + '#gid=' + blackoutGid;

    var sentCount = 0;
    var summaryLines = [];

    recipients.forEach(function (rec) {
        var name = rec.name;
        var email = rec.email;

        var subject = "SVCA Sunday School Blackout Dates";

        var plainBody =
            'Dear Co-Workers in Christ,\n\n' +

            'Praise the Lord! When you receive this email, it means that we are serving together in nurturing the next generation for the Lord and are committed to the SVCA Children’s Sunday School ministry.\n\n' +

            'Please mark the dates when you cannot serve, and the system will automatically generate a rotation schedule based on the rules, minimizing the chance of human errors in manual scheduling. Just go to https://docs.google.com/spreadsheets/d/1UmGhZH8p5cqZSktto-i2qV5PuGH607UF8UTv_VGe9C8/edit#gid=1893596443 (you must log in with your SVCA email) and check the dates you **cannot serve** on the row corresponding to your name. Please be careful not to make changes on other co-workers’ rows.\n\n' +

            'We kindly ask everyone to complete it before Feb 8 so that we will have enough time to arrange the service schedule of next quarter.\n\n' +

            'May the Lord help us improve the quality of our service together and be good stewards of the time He gives us.\n\n' +

            'In Christ,\n\n' +
            'Sister Deborah\n' +
            'Children’s Sunday School Co-Worker';

        var htmlBody =
            '親愛的同工' + name + ':' + '<br><br>' +

            '感謝主，當您收到這封Email，表示我們一起為主培育下一代，委身於SVCA兒童主日學事工。<br><br>' +

            '同工请自已把<b>“無法上崗”</b>的日期圈選出來，其他時間開放讓系統自動按照規則輪值，盡量避免人工安排時的疏漏。' +

            '請點擊 <a href="https://docs.google.com/spreadsheets/d/1UmGhZH8p5cqZSktto-i2qV5PuGH607UF8UTv_VGe9C8/edit#gid=1893596443">Blackout Dates 表格鏈接</a>（需要用您的svca email），在您名字對應的那一行勾選您<strong>無法上崗</strong>的日期。注意請不要在其他同工的行上勾選。<br><br>' +

            '敬請大家在 2/8 周日之前完成以便负责同工有足够时间排下个季度的服事时间表。<br><br>' +

            '求主幫助我們一起提升服事的品質，做時間的好管家。<br><br>' +


            '雅慧姐妹<br>' +
            '兒童主日學同工<br><br><br>' +

            'Dear Co-Workers ' + name + ' in Christ,<br><br>' +

            'Praise the Lord! When you receive this email, it means that we are serving together in nurturing the next generation for the Lord and are committed to the SVCA Children’s Sunday School ministry.<br><br>' +

            'Please mark the dates when you <strong>cannot serve</strong>, and the system will automatically generate a rotation schedule based on the rules, minimizing the chance of human errors in manual scheduling. ' +

            'Please click the <a href="https://docs.google.com/spreadsheets/d/1UmGhZH8p5cqZSktto-i2qV5PuGH607UF8UTv_VGe9C8/edit#gid=1893596443">Blackout Dates link</a> (you must log in with your SVCA email) and check the dates you <strong>cannot serve</strong> on the row corresponding to your name. Please be careful not to make changes on other co-workers’ rows.<br><br>' +

            'We kindly ask everyone to complete it before Sunday Feb 8.<br><br>' +

            'May the Lord help us improve the quality of our service together and be good stewards of the time He gives us.<br><br>' +

            'In Christ,<br><br>' +
            'Sister Deborah<br><br>' +
            'Children’s Sunday School Co-Worker<br>';

        MailApp.sendEmail({
            to: email,
            subject: subject,
            body: plainBody,
            htmlBody: htmlBody
        });

        // Log each email send
        var desc = "Sent blackout dates email to " + name + " (" + email + ")";
        if (typeof logAction === "function") {
            logAction(desc);
        }

        sentCount++;
        summaryLines.push(name + " (" + email + ")");
    });

    SpreadsheetApp.getUi().alert(
        "Sent " + sentCount + " test email(s) to:\n\n" + summaryLines.join("\n")
    );
}

/**
 * Sends email notifications to volunteers who have assignments
 * on the upcoming Sunday (first Sunday on/after today).
 *
 * The email lists the roles for that person on that date.
 */
function sendUpcomingSundayAssignmentsEmail() {
    var ss = SpreadsheetApp.getActive();
    var scheduleSheet = ss.getSheetByName('Schedule');
    var rolesSheet = ss.getSheetByName('Roles');

    if (!scheduleSheet || !rolesSheet) {
        SpreadsheetApp.getUi().alert('Missing "Schedule" or "Roles" sheet. Please check sheet names.');
        return;
    }

    // Find the upcoming Sunday date from the Schedule.
    var upcomingDate = findUpcomingSundayDate_(scheduleSheet);
    if (!upcomingDate) {
        SpreadsheetApp.getUi().alert('No upcoming Sunday found in the Schedule sheet.');
        return;
    }

    // Get all schedule data.
    var scheduleValues = scheduleSheet.getDataRange().getValues();
    var header = scheduleValues[0]; // row 1: Date | Role1 | Role2 | ...

    // Find the row for the upcoming Sunday.
    var targetRowIndex = -1;
    var upcomingTime = new Date(upcomingDate.getTime());
    upcomingTime.setHours(0, 0, 0, 0);

    for (var i = 1; i < scheduleValues.length; i++) {
        var d = scheduleValues[i][0];
        if (d instanceof Date) {
            var dClean = new Date(d.getTime());
            dClean.setHours(0, 0, 0, 0);
            if (dClean.getTime() === upcomingTime.getTime()) {
                targetRowIndex = i;
                break;
            }
        }
    }

    if (targetRowIndex === -1) {
        SpreadsheetApp.getUi().alert('Upcoming Sunday date was found, but no matching row in Schedule.');
        return;
    }

    var row = scheduleValues[targetRowIndex];

    // Build a map: volunteerName -> [roles...]
    // Skip columns whose header contains "Helper" (e.g., "xxx Helper")
    var assignmentsByPerson = {};
    for (var c = 1; c < header.length; c++) { // column 0 is Date, so start from 1
        var roleName = header[c];
        var volName = row[c];

        // Skip "Helper" roles - do not send notifications for these volunteers
        if (roleName && String(roleName).indexOf('Helper') !== -1) {
            continue;
        }

        if (volName && typeof volName === 'string') {
            if (!assignmentsByPerson[volName]) {
                assignmentsByPerson[volName] = [];
            }
            assignmentsByPerson[volName].push(roleName);
        }
    }

    if (Object.keys(assignmentsByPerson).length === 0) {
        SpreadsheetApp.getUi().alert('No assignments found for the upcoming Sunday.');
        return;
    }

    // Build a map from volunteer name -> email from the Roles sheet.
    var rolesValues = rolesSheet.getDataRange().getValues();
    var nameToEmail = {};

    if (rolesValues.length > 1) {
        // Find the email column by header name "SVCA Email"
        var headerRow = rolesValues[0];
        var emailColIndex = -1;
        for (var c = 0; c < headerRow.length; c++) {
            if (String(headerRow[c]).trim() === 'SVCA Email') {
                emailColIndex = c;
                break;
            }
        }

        if (emailColIndex === -1) {
            logAction('Warning: "SVCA Email" column not found in Roles sheet');
        } else {
            for (var r = 1; r < rolesValues.length; r++) {
                var name = rolesValues[r][0];          // Col A: Name
                var email = rolesValues[r][emailColIndex]; // "SVCA Email" column
                if (name && email) {
                    nameToEmail[name] = email;
                }
            }
        }
    }

    // Also add volunteers from the "Parent Helper" sheet if not already in the map.
    // Parent Helper: Col A = English name, Col C = Chinese name, Col D = Email
    // Full name on Schedule: "<Chinese_name> <English_Name>" or just English name if no Chinese.
    var parentHelperSheet = ss.getSheetByName('Parent Helper');
    if (parentHelperSheet) {
        var phValues = parentHelperSheet.getDataRange().getValues();
        for (var r = 1; r < phValues.length; r++) { // Start from row 2 (index 1)
            var englishName = phValues[r][0];  // Col A: English name
            var chineseName = phValues[r][2];  // Col C: Chinese name
            var phEmail = phValues[r][3];       // Col D: Email

            if (!englishName || !phEmail) continue;

            // Construct full name: "<Chinese_name> <English_Name>" or just English name
            var fullName;
            if (chineseName && String(chineseName).trim() !== '') {
                fullName = String(chineseName).trim() + ' ' + String(englishName).trim();
            } else {
                fullName = String(englishName).trim();
            }

            // Only add if not already present from Roles sheet
            if (!nameToEmail[fullName]) {
                nameToEmail[fullName] = phEmail;
            }
        }
    }

    var tz = ss.getSpreadsheetTimeZone();
    var dateStr = Utilities.formatDate(upcomingDate, tz, 'MMM d, yyyy');
    var sheetUrl = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQEee23wi5-Z1-6wth7iPc_uftGoJjsjuBnlnqExgn-K0wBlVUu01ZltHEUY89iDf6vF1kBRI2tSocU/pubhtml?gid=1967283750&single=true';

    var sentCount = 0;

    // Send one email per volunteer.
    for (var person in assignmentsByPerson) {
        var email = nameToEmail[person];
        if (!email) {
            // No email on file; skip this person.
            logAction('No email found for ' + person);
            continue;
        }
        email = email.trim();

        var rolesList = assignmentsByPerson[person];
        var subject = 'SVCA Children’s Ministry – This Sunday (' + dateStr + ')';

        var bodyLines = [];
        bodyLines.push('Hi ' + person + ',');
        bodyLines.push('');
        bodyLines.push('Here are your assignments for this Sunday (' + dateStr + '):');
        bodyLines.push('');

        for (var i = 0; i < rolesList.length; i++) {
            bodyLines.push('• ' + rolesList[i]);
        }

        bodyLines.push('');
        bodyLines.push('If you cannot serve this Sunday for some reason, please inform Sister Selena by replying this email as soon as possible. You can also view the full schedule here:');
        bodyLines.push(sheetUrl);
        bodyLines.push('');
        bodyLines.push('Thank you for serving!');
        bodyLines.push('');
        bodyLines.push('SVCA Children’s Ministry');

        var body = bodyLines.join('\n');

        //MailApp.sendEmail(email, subject, body);
        MailApp.sendEmail({
            to: email,
            replyTo: 'shuru.fang@svca.cc',
            subject: subject,
            body: body
        });

        logAction('Sent weekly notification to ' + person);
        sentCount++;
    }

    SpreadsheetApp.getUi().alert('Sent ' + sentCount + ' emails for ' + dateStr + '.');
}
