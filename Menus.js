/**
 * Adds a custom menu when the spreadsheet is opened.
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var editors = ss.getEditors().map(function (e) {
        return e.getEmail().toLowerCase();
    });
    var user = Session.getActiveUser().getEmail().toLowerCase();
    var isEditor = editors.indexOf(user) !== -1;

    var menu = ui.createMenu("Service Scheduler");

    if (isEditor) {
        // Full admin menu
        menu
            .addItem("Refresh Blackout Dates", "refreshBlackoutDates")
            .addItem("Lock Blackout Dates", "lockBlackoutDates")
            .addItem("Unlock Blackout Dates", "unlockBlackoutDates")
            .addSeparator()
            //.addItem("Auto Populate Schedule", "autoPopulateSchedule")
            .addItem("Highlight Conflicts", "highlightConflicts")
            .addItem("Highlight One Person", "highlightOnePerson")
            .addSeparator()
            .addItem("Copy to Schedule History", "copyScheduleToHistory")
            .addSeparator()
            .addItem("Send Blackout Notification Emails", "sendBlackoutNotificationEmails")
            .addItem("Send Upcoming Sunday Emails", "sendUpcomingSundayAssignmentsEmail")
            .addToUi();
    } else {
        // Non-editors see an *empty* (or minimal) menu
        // You can choose one:

        // 1) A completely hidden menu (NO menu items):
        // (Do not add any items)

        // 2) Or a minimal menu with just one help item:
        // .addItem("About Scheduler", "showHelpMessage");
    }

    menu.addToUi();
}