function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui
    .createMenu('Tester Automation')
    .addItem('Initialize Spreadsheet', 'menuInitializeSpreadsheet')
    .addItem('Run Health Check', 'menuRunHealthCheck')
    .addSeparator()
    .addItem('Run Tester Smoke Test', 'menuRunTesterSmokeTest')
    .addSeparator()
    .addItem('Seed Default Tasks', 'menuSeedDefaultTasks')
    .addItem('Assign Today\'s Tasks for Active Testers', 'menuAssignTodayTasks')
    .addItem('Run Task Engine Smoke Test', 'menuRunTaskEngineSmokeTest')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Email Automation')
        .addItem('Run Morning Reminders', 'menuRunMorningReminders')
        .addItem('Run Evening Reminders', 'menuRunEveningReminders')
        .addSeparator()
        .addItem('Install Reminder Triggers', 'menuInstallReminderTriggers')
        .addItem('Remove Reminder Triggers', 'menuRemoveReminderTriggers')
        .addItem('Run Email Automation Smoke Test', 'menuRunEmailAutomationSmokeTest')
    )
    .addSubMenu(
      ui
        .createMenu('Feedback')
        .addItem('Configure Web App URL', 'menuConfigureFeedbackWebAppUrl')
        .addItem('Check Web App Configuration', 'menuCheckFeedbackConfiguration')
        .addItem('Run Feedback Smoke Test', 'menuRunFeedbackSmokeTest')
        .addItem('Show Deployment Instructions', 'menuShowFeedbackDeploymentInstructions')
    )
    .addSubMenu(
      ui
        .createMenu('Monitoring')
        .addItem('Refresh Monitoring', 'menuRunMonitoringRefresh')
        .addItem('Refresh Tester Statuses', 'menuRefreshTesterStatuses')
        .addSeparator()
        .addItem('Install Monitoring Trigger', 'menuInstallMonitoringTrigger')
        .addItem('Remove Monitoring Trigger', 'menuRemoveMonitoringTrigger')
        .addItem('Run Monitoring Smoke Test', 'menuRunMonitoringSmokeTest')
    )
    .addToUi();
}

function menuInitializeSpreadsheet() {
  runMenuAction_(function () {
    var result = initializeSpreadsheet();
    showMenuToast_(
      'Initialization complete\nCreated sheets: ' + result.createdSheets.length +
        '\nExisting sheets: ' + result.existingSheets.length,
      'Tester Automation'
    );
  });
}

function menuRunHealthCheck() {
  runMenuAction_(function () {
    var result = healthCheck();
    if (!result.ok) {
      showMenuAlert_(
        'Health Check Failed',
        'Missing sheets: ' + result.missingSheets.join(', ')
      );
      return;
    }
    showMenuToast_(
      'Health check passed\nAll ' + result.requiredSheets.length + ' required sheets exist.',
      'Tester Automation'
    );
  });
}

function menuRunTesterSmokeTest() {
  runMenuAction_(function () {
    var result = runTesterServiceSmokeTest();
    showMenuToast_(
      'Tester smoke test passed\nStatus: ' + result.status,
      'Tester Automation'
    );
  });
}

function menuSeedDefaultTasks() {
  runMenuAction_(function () {
    var result = seedDefaultTasks();
    showMenuToast_(
      'Task seeding complete\nCreated: ' + result.created + '\nSkipped: ' + result.skipped,
      'Tester Automation'
    );
  });
}

function menuAssignTodayTasks() {
  runMenuAction_(function () {
    var result = assignTodayTasksForAllActiveTesters();
    showMenuToast_(
      'Today\'s task assignment complete\nProcessed: ' + result.processed +
        '\nAssigned: ' + result.assigned +
        '\nExisting: ' + result.existing +
        '\nFailed: ' + result.errors.length,
      'Tester Automation'
    );
  });
}

function menuRunTaskEngineSmokeTest() {
  runMenuAction_(function () {
    var result = runTaskEngineSmokeTest();
    showMenuToast_(
      'Task engine smoke test passed\nDay: ' + result.day +
        '\nTasks: ' + result.taskCount,
      'Tester Automation'
    );
  });
}

function menuRunMorningReminders() {
  runMenuAction_(function () {
    var result = sendMorningReminders();
    showReminderSummary_('Morning reminders complete', result);
  });
}

function menuRunEveningReminders() {
  runMenuAction_(function () {
    var result = sendEveningReminders();
    showReminderSummary_('Evening reminders complete', result);
  });
}

function menuRunEmailAutomationSmokeTest() {
  runMenuAction_(function () {
    var result = runEmailAutomationSmokeTest();
    showMenuToast_(
      'Email automation smoke test passed\nMorning sent: ' + result.morningSent +
        '\nEvening sent: ' + result.eveningSent,
      'Tester Automation'
    );
  });
}

function menuInstallReminderTriggers() {
  runMenuAction_(function () {
    var result = installReminderTriggers();
    showMenuToast_(
      'Reminder triggers ready\nMorning created: ' + result.createdMorning +
        '\nEvening created: ' + result.createdEvening,
      'Tester Automation'
    );
  });
}

function menuRemoveReminderTriggers() {
  runMenuAction_(function () {
    var result = removeReminderTriggers();
    showMenuToast_(
      'Reminder triggers removed: ' + result.removed,
      'Tester Automation'
    );
  });
}

function menuCheckFeedbackConfiguration() {
  runMenuAction_(function () {
    getFeedbackWebAppUrl();
    showMenuToast_(
      'Feedback Web App URL is configured.\nEmail Test Mode: ' + getEmailTestMode(),
      'Tester Automation'
    );
  });
}

function menuConfigureFeedbackWebAppUrl() {
  runMenuAction_(function () {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      'Configure Feedback Web App',
      'Paste the Web App deployment URL ending in /exec:',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      showMenuToast_('Web App URL was not changed.', 'Tester Automation');
      return;
    }

    setFeedbackWebAppUrl(response.getResponseText());
    showMenuToast_('Feedback Web App URL saved.', 'Tester Automation');
  });
}

function menuRunFeedbackSmokeTest() {
  runMenuAction_(function () {
    var result = runFeedbackSmokeTest();
    showMenuAlert_(
      'Feedback Smoke Test Passed',
      'Feedback smoke test passed\nDuplicate blocked: ' + result.duplicateBlocked +
        '\nActivity completed: ' + result.activityCompleted +
        '\n\nOpen this fresh test form URL:\n' + result.testFormUrl
    );
  });
}

function menuShowFeedbackDeploymentInstructions() {
  showMenuAlert_(
    'Feedback Web App Deployment',
    '1. Open Apps Script.\n' +
      '2. Select Deploy > New deployment.\n' +
      '3. Choose Web app.\n' +
      '4. Execute as the script owner.\n' +
      '5. Allow access for intended external testers.\n' +
      '6. Copy the /exec deployment URL.\n' +
      '7. Run setFeedbackWebAppUrl("DEPLOYMENT_URL").'
  );
}

function menuRunMonitoringRefresh() {
  runMenuAction_(function () {
    var result = runMonitoringRefresh();
    showMonitoringSummary_('Monitoring refresh complete', result);
  });
}

function menuRefreshTesterStatuses() {
  runMenuAction_(function () {
    var result = refreshAllTesterStatuses();
    showMonitoringSummary_('Tester status refresh complete', result);
  });
}

function menuInstallMonitoringTrigger() {
  runMenuAction_(function () {
    var result = installMonitoringTrigger();
    showMenuToast_(
      result.created
        ? 'Monitoring trigger installed for approximately 7:00 PM.'
        : 'Monitoring trigger already exists.',
      'Tester Automation'
    );
  });
}

function menuRemoveMonitoringTrigger() {
  runMenuAction_(function () {
    var result = removeMonitoringTrigger();
    showMenuToast_(
      'Monitoring triggers removed: ' + result.removed,
      'Tester Automation'
    );
  });
}

function menuRunMonitoringSmokeTest() {
  runMenuAction_(function () {
    var result = runMonitoringSmokeTest();
    showMenuAlert_(
      'Monitoring Smoke Test Passed',
      'Scenarios verified: ' + result.scenarioCount +
        '\nIdempotent refresh: ' + result.idempotent +
        '\nInactive preserved: ' + result.inactivePreserved +
        '\nMonitoring row generated: ' + result.monitoringRowGenerated
    );
  });
}

function runMenuAction_(action) {
  try {
    action();
  } catch (error) {
    showMenuAlert_(
      'Tester Automation Error',
      error && error.message ? error.message : 'The operation could not be completed.'
    );
  }
}

function showReminderSummary_(heading, result) {
  var message =
    heading +
    '\nProcessed: ' + result.processed +
    '\nSent: ' + result.sent +
    '\nSkipped: ' + result.skipped +
    '\nFailed: ' + result.failed;
  if (result.failed > 0) {
    if (result.errors.length > 0) {
      message += '\nFirst error: ' + result.errors[0].message;
    }
    showMenuAlert_('Tester Automation', message);
    return;
  }
  showMenuToast_(message, 'Tester Automation');
}

function showMonitoringSummary_(heading, result) {
  var counts = result.statusCounts || {};
  var message =
    heading +
    '\nProcessed: ' + result.processed +
    '\nStatus changes: ' + result.changed +
    '\nNot Started: ' + (counts[TESTER_STATUSES.NOT_STARTED] || 0) +
    '\nActive: ' + (counts[TESTER_STATUSES.ACTIVE] || 0) +
    '\nNeeds Reminder: ' + (counts[TESTER_STATUSES.NEEDS_REMINDER] || 0) +
    '\nAt Risk: ' + (counts[TESTER_STATUSES.AT_RISK] || 0) +
    '\nCompleted: ' + (counts[TESTER_STATUSES.COMPLETED] || 0) +
    '\nInactive: ' + (counts[TESTER_STATUSES.INACTIVE] || 0) +
    '\nFailed: ' + result.failed;
  if (result.failed > 0) {
    if (result.errors && result.errors.length > 0) {
      message += '\nFirst error: ' + result.errors[0].message;
    }
    showMenuAlert_('Tester Automation', message);
    return;
  }
  showMenuToast_(message, 'Tester Automation');
}

function showMenuToast_(message, title) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 8);
}

function showMenuAlert_(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function initializeSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var result = {
    createdSheets: [],
    initializedHeaders: [],
    existingSheets: []
  };

  Object.keys(REQUIRED_SHEET_HEADERS).forEach(function (sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      result.createdSheets.push(sheetName);
    } else {
      result.existingSheets.push(sheetName);
    }

    var headers = REQUIRED_SHEET_HEADERS[sheetName];
    if (headers.length > 0 && sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      result.initializedHeaders.push(sheetName);
    }
  });

  if (typeof ensureActivityLogReminderColumns_ === 'function') {
    ensureActivityLogReminderColumns_();
  }

  Logger.log(JSON.stringify(result));
  return result;
}

function healthCheck() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var requiredSheets = Object.keys(REQUIRED_SHEET_HEADERS);
  var missingSheets = requiredSheets.filter(function (sheetName) {
    return spreadsheet.getSheetByName(sheetName) === null;
  });
  var result = {
    ok: missingSheets.length === 0,
    spreadsheetName: spreadsheet.getName(),
    requiredSheets: requiredSheets,
    missingSheets: missingSheets
  };

  Logger.log(JSON.stringify(result));
  return result;
}
