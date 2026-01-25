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

    var lastCol = rolesSheet.getLastColumn();
    var ranges = rangeList.getRanges();
    var recipients = [];  // {name, email, row}

    ranges.forEach(function (range) {
        var startRow = range.getRow();
        var endRow = range.getLastRow();

        for (var r = startRow; r <= endRow; r++) {
            // Skip header row
            if (r < 2) continue;

            var name = rolesSheet.getRange(r, 1).getDisplayValue().trim();        // Col A
            var email = rolesSheet.getRange(r, lastCol).getDisplayValue().trim();  // Last col = Email

            if (!name || !email) continue;

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

            'To make the overall scheduling process smoother, the Children’s Ministry plans to prepare the Sunday serving rotations quarterly in advance (Dec–Feb, Mar–May, Jun–Aug, Sep–Nov). Each area coordinator will then make adjustments according to actual needs. To support this process, we have created a Blackout Dates form. Please mark the dates when you cannot serve, and the system will automatically generate a rotation schedule based on the rules, minimizing the chance of human errors in manual scheduling.\n\n' +

            'Please go to https://docs.google.com/spreadsheets/d/1UmGhZH8p5cqZSktto-i2qV5PuGH607UF8UTv_VGe9C8/edit#gid=1893596443 (you must log in with your SVCA email) and check the dates you **cannot serve** on the row corresponding to your name. Please be careful not to make changes on other co-workers’ rows.\n\n' +

            'This is the first time we are opening this scheduling process. We kindly ask everyone to complete it before November 23. We also warmly welcome any suggestions you may have—please write them in the Improvement Ideas sheet. We will do our best to continually improve this system.\n\n' +

            'May the Lord help us improve the quality of our service together and be good stewards of the time He gives us.\n\n' +

            'In Christ,\n\n' +
            'Sister Deborah\n' +
            'Children’s Sunday School Co-Worker';

        var htmlBody =
            '親愛的同工' + name + ':' + '<br><br>' +

            '感謝主，當您收到這張排班表，表示我們一起為主培育下一代，委身於SVCA兒童主日學事工。<br><br>' +

            '為了整體排班上更順暢，兒童部擬將主日服事輪值表以季度方式（12-2, 3-5, 6-8, 9-11月)預先總體安排，屆時由各項負責同工按照實際情況調動，特別設計了 Blackout Dates 表格。同工自已把<b>“無法上崗”</b>的日期圈選出來，其他時間開放讓系統自動按照規則輪值，盡量避免人工安排時的疏漏。<br><br>' +

            '請點擊 <a href="https://docs.google.com/spreadsheets/d/1UmGhZH8p5cqZSktto-i2qV5PuGH607UF8UTv_VGe9C8/edit#gid=1893596443">Blackout Dates 表格鏈接</a>（需要用您的svca email），在您名字對應的那一行勾選您<strong>無法上崗</strong>的日期。注意請不要在其他同工的行上勾選。<br><br>' +

            '這是第一次開放排班，敬請大家在 11/23 周日之前完成，所有改進意見也非常歡迎填寫在 Improvement Ideas 表格上，我們盡力將這個系統不斷完善。<br><br>' +

            '求主幫助我們一起提升服事的品質，做時間的好管家。<br><br>' +


            '雅慧姐妹<br>' +
            '兒童主日學同工<br><br><br>' +

            'Dear Co-Workers ' + name + ' in Christ,<br><br>' +

            'Praise the Lord! When you receive this email, it means that we are serving together in nurturing the next generation for the Lord and are committed to the SVCA Children’s Sunday School ministry.<br><br>' +

            'To make the overall scheduling process smoother, the Children’s Ministry plans to prepare the Sunday serving rotations quarterly in advance (Dec–Feb, Mar–May, Jun–Aug, Sep–Nov). Each area coordinator will then make adjustments according to actual needs. To support this process, we have created a Blackout Dates form. Please mark the dates when you <strong>cannot serve</strong>, and the system will automatically generate a rotation schedule based on the rules, minimizing the chance of human errors in manual scheduling.<br><br>' +

            'Please click the <a href="https://docs.google.com/spreadsheets/d/1UmGhZH8p5cqZSktto-i2qV5PuGH607UF8UTv_VGe9C8/edit#gid=1893596443">Blackout Dates link</a> (you must log in with your SVCA email) and check the dates you <strong>cannot serve</strong> on the row corresponding to your name. Please be careful not to make changes on other co-workers’ rows.<br><br>' +

            'This is the first time we are opening this scheduling process. We kindly ask everyone to complete it before November 23. We also warmly welcome any suggestions you may have—please write them in the Improvement Ideas sheet. We will do our best to continually improve this system.<br><br>' +

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
    var assignmentsByPerson = {};
    for (var c = 1; c < header.length; c++) { // column 0 is Date, so start from 1
        var roleName = header[c];
        var volName = row[c];

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
