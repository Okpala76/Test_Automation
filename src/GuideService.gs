function openGuide() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(GUIDE_SHEET_NAME);
  if (!sheet) {
    throw new Error('The Guide sheet is missing. Run Initialize Spreadsheet first.');
  }
  spreadsheet.setActiveSheet(sheet);
  return sheet;
}

function refreshGuide() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    return refreshGuideWithoutLock_();
  } finally {
    lock.releaseLock();
  }
}

function refreshGuideWithoutLock_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUIDE_SHEET_NAME);
  if (!sheet) {
    throw new Error('The Guide sheet is missing. Run Initialize Spreadsheet first.');
  }

  var requiredRows = 74;
  var requiredColumns = 6;
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      requiredColumns - sheet.getMaxColumns()
    );
  }

  sheet
    .getRange(1, 1, sheet.getMaxRows(), requiredColumns)
    .breakApart();
  sheet.clear();
  sheet.setConditionalFormatRules([]);
  sheet.getCharts().forEach(function (chart) {
    sheet.removeChart(chart);
  });

  sheet
    .getRange('A1:F1')
    .merge()
    .setValue('Tester Automation Quick-Start Guide');
  sheet
    .getRange('A2:F2')
    .merge()
    .setValue('A concise operating guide for the 14-day tester program.');

  writeGuideSection_(sheet, 4, 'WHAT THIS SPREADSHEET DOES');
  writeGuideParagraph_(
    sheet,
    5,
    'Coordinates tester enrollment, daily activities, independent email and WhatsApp reminders, feedback, participation monitoring, and progress reporting in one place.'
  );

  writeGuideSection_(sheet, 7, 'THE 14-DAY TESTING SYSTEM');
  writeGuideParagraph_(
    sheet,
    8,
    'Each tester follows a 14-day plan with one morning activity and one evening check. Activity and feedback records show whether the test is progressing as expected.'
  );

  writeGuideSection_(sheet, 10, 'REAL TESTER ONBOARDING');
  writeGuideTableHeader_(sheet, 11, 'Step', 'Action');
  sheet.getRange(12, 1, 5, 2).setValues([
    ['1', "Collect each tester's Name and Google Play email."],
    ['2', 'Paste only Name and Email into blank rows in the Testers sheet.'],
    ['3', 'Run Tester Automation > Testers > Import New Testers. The system fills Tester ID, Status, Token, Created At, and Updated At.'],
    ['4', 'When the closed-test start date is confirmed, run Tester Automation > Testers > Activate Not Started Testers. Enter the common Start Date; the system sets Start Date and Status = Active.'],
    ['5', 'The existing morning and evening automation then takes over.']
  ]);
  writeGuideParagraph_(
    sheet,
    17,
    'Important: Do not activate testers until the actual testing start date is confirmed.'
  );

  writeGuideSection_(sheet, 19, 'ADDING A FRESH TESTER DURING A LIVE TEST');
  writeGuideTableHeader_(sheet, 20, 'Step', 'Action');
  sheet.getRange(21, 1, 8, 2).setValues([
    ['1', "Add the tester's Google account email to Play Console."],
    ['2', 'Have them opt in and install the app.'],
    ['3', 'Add Name + Email to a blank Testers row.'],
    ['4', 'Run Testers > Import New Testers.'],
    ['5', "Click any cell on that tester's row."],
    ['6', 'Run Testers > Activate Selected Tester.'],
    ['7', 'Enter their real start date.'],
    ['8', 'Their Day 1-14 timeline now runs independently.']
  ]);

  writeGuideSection_(sheet, 30, 'WHATSAPP REMINDERS');
  writeGuideParagraph_(
    sheet,
    31,
    'Email remains the primary reminder channel. WhatsApp is optional and is attempted independently, so a WhatsApp failure does not stop email reminders.'
  );
  writeGuideParagraph_(
    sheet,
    33,
    'A tester receives WhatsApp only when Phone is configured in international format and WhatsApp Enabled is TRUE. Missing phones and disabled testers are skipped.'
  );
  writeGuideParagraph_(
    sheet,
    35,
    'Use Tester Automation > WhatsApp to configure Evolution API, set the test recipient, enable Test Mode, check configuration, and send a harmless test message.'
  );

  writeGuideSection_(sheet, 38, 'SHEETS AT A GLANCE');
  writeGuideTableHeader_(sheet, 39, 'Sheet', 'Purpose');
  sheet.getRange(40, 1, 7, 2).setValues([
    ['Testers', 'Tester details, start dates, and current participation status.'],
    ['Tasks', 'The Day 1-14 morning and evening testing plan.'],
    ['Activity Log', 'Assigned activities, completion, and separate email/WhatsApp reminder history.'],
    ['Feedback', 'Tester ratings, completion answers, bug flags, and comments.'],
    ['Monitoring', 'Calculated participation details used for operational follow-up.'],
    ['Dashboard', 'A summarized view of status, activity, feedback, and progress.'],
    ['Guide', 'This quick-start reference for spreadsheet operators.']
  ]);

  writeGuideSection_(sheet, 48, 'BASIC OPERATING FLOW');
  writeGuideTableHeader_(sheet, 49, 'Step', 'Action');
  sheet.getRange(50, 1, 7, 2).setValues([
    ['1', 'Paste Name and Google Play email, then import the tester.'],
    ['2', 'Activate one selected tester or all Not Started testers with confirmed Start Dates.'],
    ['3', 'The system assigns the appropriate daily tasks.'],
    ['4', 'Morning and evening email and optional WhatsApp reminders are attempted independently.'],
    ['5', 'The tester submits feedback through the feedback form.'],
    ['6', 'Monitoring calculates participation status from activity and feedback.'],
    ['7', 'Dashboard summarizes overall progress and attention required.']
  ]);

  writeGuideSection_(sheet, 58, 'STATUS MEANINGS');
  writeGuideTableHeader_(sheet, 59, 'Status', 'Meaning');
  sheet.getRange(60, 1, 6, 2).setValues([
    [TESTER_STATUSES.NOT_STARTED, 'Testing has not begun or the Start Date is in the future.'],
    [TESTER_STATUSES.ACTIVE, 'The tester is participating recently or is in the initial grace period.'],
    [TESTER_STATUSES.NEEDS_REMINDER, 'No meaningful participation has been recorded for 2 days.'],
    [TESTER_STATUSES.AT_RISK, 'No meaningful participation has been recorded for 3 or more days.'],
    [TESTER_STATUSES.COMPLETED, 'The testing period ended with recorded participation.'],
    [TESTER_STATUSES.INACTIVE, 'The tester was manually excluded from automation.']
  ]);

  writeGuideSection_(sheet, 67, 'TEST MODE AND MANUAL CONTROLS');
  writeGuideParagraph_(
    sheet,
    68,
    'Email and WhatsApp have separate Test Modes and separate test recipients. WhatsApp Test Mode never falls back to a real tester phone.'
  );
  writeGuideParagraph_(
    sheet,
    70,
    'Before live WhatsApp delivery, run Check WhatsApp Configuration and Send WhatsApp Test Message while WhatsApp Test Mode is enabled.'
  );
  writeGuideParagraph_(
    sheet,
    72,
    'Run initialization, reminders, monitoring, dashboard refreshes, readiness checks, and other manual controls from the Tester Automation menu.'
  );

  formatGuide_(sheet);
  return {
    sheetName: GUIDE_SHEET_NAME,
    sectionCount: 9
  };
}

function writeGuideSection_(sheet, row, title) {
  sheet.getRange(row, 1, 1, 6).merge().setValue(title);
}

function writeGuideParagraph_(sheet, row, text) {
  sheet.getRange(row, 1, 1, 6).merge().setValue(text);
}

function writeGuideTableHeader_(sheet, row, firstHeader, secondHeader) {
  sheet.getRange(row, 1).setValue(firstHeader);
  sheet.getRange(row, 2, 1, 5).merge().setValue(secondHeader);
}

function formatGuide_(sheet) {
  var darkGreen = '#214d3b';
  var sectionGreen = '#d9ead3';
  var headerFill = '#edf3ef';
  var border = '#c8d2ca';

  sheet
    .getRange('A1:F1')
    .setBackground(darkGreen)
    .setFontColor('#ffffff')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('A2:F2')
    .setBackground('#f4f7f5')
    .setFontColor('#516158')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);

  [4, 7, 10, 19, 30, 38, 48, 58, 67].forEach(function (row) {
    sheet
      .getRange(row, 1, 1, 6)
      .setBackground(sectionGreen)
      .setFontColor('#17382b')
      .setFontWeight('bold');
  });
  [11, 20, 39, 49, 59].forEach(function (row) {
    sheet
      .getRange(row, 1, 1, 6)
      .setBackground(headerFill)
      .setFontWeight('bold');
  });

  sheet.getRange(11, 1, 6, 6).setBorder(
    true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(20, 1, 9, 6).setBorder(
    true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(39, 1, 8, 6).setBorder(
    true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(49, 1, 8, 6).setBorder(
    true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(59, 1, 7, 6).setBorder(
    true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
  );

  [
    { startRow: 12, rowCount: 5 },
    { startRow: 21, rowCount: 8 },
    { startRow: 40, rowCount: 7 },
    { startRow: 50, rowCount: 7 },
    { startRow: 60, rowCount: 6 }
  ].forEach(function (table) {
    for (var row = table.startRow; row < table.startRow + table.rowCount; row += 1) {
      sheet.getRange(row, 2, 1, 5).merge();
    }
  });

  sheet.getRange(1, 1, requiredGuideLastRow_(), 6).setWrap(true);
  sheet.getRange(1, 1, requiredGuideLastRow_(), 6).setVerticalAlignment('middle');
  sheet.getRange('A17:F17').setBackground('#fff2cc').setFontWeight('bold');
  sheet.getRange('A68:F68').setBackground('#fff2cc').setFontWeight('bold');
  sheet.setColumnWidth(1, 175);
  [2, 3, 4, 5, 6].forEach(function (column) {
    sheet.setColumnWidth(column, 125);
  });
  sheet.setRowHeights(12, 5, 50);
  sheet.setRowHeights(21, 8, 42);
  sheet.setRowHeights(40, 7, 38);
  sheet.setRowHeights(50, 7, 42);
  sheet.setRowHeights(60, 6, 38);
  sheet.setRowHeight(5, 42);
  sheet.setRowHeight(8, 46);
  sheet.setRowHeight(14, 62);
  sheet.setRowHeight(15, 70);
  sheet.setRowHeight(17, 42);
  sheet.setRowHeight(31, 50);
  sheet.setRowHeight(33, 50);
  sheet.setRowHeight(35, 50);
  sheet.setRowHeight(68, 50);
  sheet.setRowHeight(70, 50);
  sheet.setRowHeight(72, 44);
  sheet.setFrozenRows(2);
}

function requiredGuideLastRow_() {
  return 74;
}
