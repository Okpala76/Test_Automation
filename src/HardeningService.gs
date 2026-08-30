function getSystemConfigurationHealth() {
  var properties = PropertiesService.getScriptProperties();
  var emailTestMode = properties.getProperty(EMAIL_TEST_MODE_PROPERTY) === 'true';
  var testRecipient = properties.getProperty(TEST_EMAIL_RECIPIENT_PROPERTY) || '';
  var webAppUrl = properties.getProperty(FEEDBACK_WEB_APP_URL_PROPERTY) || '';
  var recipientValid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(testRecipient.trim());
  var webAppUrlValid = /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(
    webAppUrl.trim()
  );
  var warnings = [];
  var failures = [];

  if (!webAppUrl) {
    failures.push('Feedback Web App URL is not configured.');
  } else if (!webAppUrlValid) {
    failures.push('Feedback Web App URL is invalid.');
  }
  if (!testRecipient) {
    warnings.push('Test email recipient is not configured.');
  } else if (!recipientValid) {
    warnings.push('Test email recipient is invalid.');
  }
  if (emailTestMode) {
    warnings.push('Email Test Mode is enabled; live tester delivery is disabled.');
  }

  return {
    ok: failures.length === 0,
    emailTestMode: emailTestMode,
    testRecipientConfigured: Boolean(testRecipient),
    testRecipientValid: recipientValid,
    feedbackWebAppUrlConfigured: Boolean(webAppUrl),
    feedbackWebAppUrlValid: webAppUrlValid,
    scriptTimeZone: Session.getScriptTimeZone(),
    safeForDryRun: emailTestMode && recipientValid && webAppUrlValid,
    readyForLiveSending: !emailTestMode && webAppUrlValid,
    failures: failures,
    warnings: warnings
  };
}

function getAutomationTriggerHealth() {
  var specs = getAutomationTriggerSpecs_();
  var triggers = ScriptApp.getProjectTriggers();
  var matchedCount = 0;
  var handlers = specs.map(function (spec) {
    var matching = triggers.filter(function (trigger) {
      return trigger.getHandlerFunction() === spec.handler;
    });
    var valid = matching.filter(isValidClockTrigger_);
    matchedCount += matching.length;
    return {
      handler: spec.handler,
      expectedHour: spec.hour,
      count: matching.length,
      validClockCount: valid.length,
      missing: valid.length === 0,
      duplicate: matching.length > 1,
      invalidType: matching.length !== valid.length,
      healthy: matching.length === 1 && valid.length === 1
    };
  });

  return {
    ok: handlers.every(function (handler) {
      return handler.healthy;
    }),
    handlers: handlers,
    scope: 'current_user',
    managedTriggerCount: matchedCount,
    unrelatedTriggerCount: triggers.length - matchedCount
  };
}

function repairAutomationTriggers() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var specs = getAutomationTriggerSpecs_();
    var triggers = ScriptApp.getProjectTriggers();
    var result = {
      created: [],
      removed: [],
      preserved: []
    };

    specs.forEach(function (spec) {
      var matching = triggers.filter(function (trigger) {
        return trigger.getHandlerFunction() === spec.handler;
      });
      var keeper = null;

      matching.forEach(function (trigger) {
        if (!keeper && isValidClockTrigger_(trigger)) {
          keeper = trigger;
          return;
        }
        ScriptApp.deleteTrigger(trigger);
        result.removed.push(spec.handler);
      });

      if (!keeper) {
        createDailyAutomationTrigger_(spec.handler, spec.hour);
        result.created.push(spec.handler);
      } else {
        result.preserved.push(spec.handler);
      }
    });

    result.health = getAutomationTriggerHealth();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function runSystemReadinessCheck() {
  var report = {
    ok: true,
    checkedAt: new Date(),
    failures: [],
    warnings: [],
    checks: []
  };
  var sheetInspection = inspectRequiredSheetStructure_();

  sheetInspection.results.forEach(function (result) {
    addReadinessResult_(
      report,
      'Sheet: ' + result.sheetName,
      result.ok ? 'pass' : 'fail',
      result.message
    );
  });

  if (sheetInspection.validSheets[TASK_SHEET_NAME]) {
    inspectDefaultTaskCoverage_().forEach(function (result) {
      addReadinessResult_(report, result.name, result.status, result.message);
    });
  }

  var configuration = getSystemConfigurationHealth();
  configuration.failures.forEach(function (message) {
    addReadinessResult_(report, 'Configuration', 'fail', message);
  });
  configuration.warnings.forEach(function (message) {
    addReadinessResult_(report, 'Configuration', 'warning', message);
  });
  if (configuration.failures.length === 0 && configuration.warnings.length === 0) {
    addReadinessResult_(report, 'Configuration', 'pass', 'Required configuration is ready.');
  }

  if (sheetInspection.validSheets[TESTER_SHEET_NAME]) {
    inspectWhatsAppReadiness_().forEach(function (result) {
      addReadinessResult_(report, result.name, result.status, result.message);
    });
  }

  try {
    var triggerHealth = getAutomationTriggerHealth();
    triggerHealth.handlers.forEach(function (handler) {
      var message = handler.healthy
        ? handler.handler + ' has exactly one time-driven trigger.'
        : handler.handler + ' requires trigger repair (found ' + handler.count + ').';
      addReadinessResult_(
        report,
        'Trigger: ' + handler.handler,
        handler.healthy ? 'pass' : 'fail',
        message
      );
    });
    addReadinessResult_(
      report,
      'Automation trigger scope',
      'warning',
      'Apps Script exposes only triggers owned by the current user; verify other editors have not installed duplicates.'
    );
  } catch (error) {
    addReadinessResult_(
      report,
      'Automation triggers',
      'fail',
      'Trigger health could not be inspected: ' + error.message
    );
  }

  report.ok = report.failures.length === 0;
  report.summary = {
    passed: report.checks.filter(function (check) {
      return check.status === 'pass';
    }).length,
    warnings: report.warnings.length,
    failures: report.failures.length
  };
  return report;
}

function cleanupSmokeTestData() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var testerSheet = spreadsheet.getSheetByName(TESTER_SHEET_NAME);
    if (!testerSheet) {
      throw new Error('The Testers sheet is missing. Nothing was deleted.');
    }
    var testerHeaders = validateCleanupSheetHeaders_(testerSheet, TESTER_SHEET_NAME);
    var lastTesterRow = testerSheet.getLastRow();
    var testerRows = lastTesterRow <= 1
      ? []
      : testerSheet
          .getRange(2, 1, lastTesterRow - 1, testerHeaders.length)
          .getValues();
    var testerIdIndex = testerHeaders.indexOf('Tester ID');
    var testerNameIndex = testerHeaders.indexOf('Name');
    var classifications = Object.create(null);

    testerRows.forEach(function (row) {
      var testerId = String(row[testerIdIndex] || '').trim();
      if (!testerId) {
        return;
      }
      if (!classifications[testerId]) {
        classifications[testerId] = { smoke: 0, nonSmoke: 0 };
      }
      if (isSmokeTestName_(row[testerNameIndex])) {
        classifications[testerId].smoke += 1;
      } else {
        classifications[testerId].nonSmoke += 1;
      }
    });

    var eligibleIds = Object.create(null);
    Object.keys(classifications).forEach(function (testerId) {
      var classification = classifications[testerId];
      if (classification.smoke > 0 && classification.nonSmoke === 0) {
        eligibleIds[testerId] = true;
      }
    });

    var feedbackSheet = spreadsheet.getSheetByName(FEEDBACK_SHEET_NAME);
    var activitySheet = spreadsheet.getSheetByName(ACTIVITY_LOG_SHEET_NAME);
    var monitoringSheet = spreadsheet.getSheetByName(MONITORING_SHEET_NAME);
    validateCleanupSheetHeaders_(feedbackSheet, FEEDBACK_SHEET_NAME);
    validateCleanupSheetHeaders_(activitySheet, ACTIVITY_LOG_SHEET_NAME);
    validateCleanupSheetHeaders_(monitoringSheet, MONITORING_SHEET_NAME);

    var deleted = {
      feedback: deleteCleanupRowsByTesterId_(
        feedbackSheet,
        FEEDBACK_SHEET_NAME,
        eligibleIds
      ),
      activities: deleteCleanupRowsByTesterId_(
        activitySheet,
        ACTIVITY_LOG_SHEET_NAME,
        eligibleIds
      ),
      monitoring: deleteCleanupRowsByTesterId_(
        monitoringSheet,
        MONITORING_SHEET_NAME,
        eligibleIds
      ),
      testers: 0
    };

    for (var rowIndex = testerRows.length - 1; rowIndex >= 0; rowIndex -= 1) {
      var testerId = String(testerRows[rowIndex][testerIdIndex] || '').trim();
      if (
        Object.prototype.hasOwnProperty.call(eligibleIds, testerId) &&
        isSmokeTestName_(testerRows[rowIndex][testerNameIndex])
      ) {
        testerSheet.deleteRow(rowIndex + 2);
        deleted.testers += 1;
      }
    }

    return {
      ok: true,
      deleted: deleted,
      eligibleTesterIds: Object.keys(eligibleIds).length,
      skippedAmbiguousTesterRows: testerRows.filter(function (row) {
        var testerId = String(row[testerIdIndex] || '').trim();
        return (
          isSmokeTestName_(row[testerNameIndex]) &&
          !Object.prototype.hasOwnProperty.call(eligibleIds, testerId)
        );
      }).length,
      tasksDeleted: 0
    };
  } finally {
    lock.releaseLock();
  }
}

function runDeterministicLogicTests() {
  var start = new Date(Date.UTC(2026, 0, 1, 12));
  var statusScenarios = [
    {
      facts: { currentStatus: TESTER_STATUSES.INACTIVE },
      expected: TESTER_STATUSES.INACTIVE
    },
    {
      facts: { currentStatus: TESTER_STATUSES.ACTIVE, hasStartDate: false },
      expected: TESTER_STATUSES.NOT_STARTED
    },
    {
      facts: {
        currentStatus: TESTER_STATUSES.ACTIVE,
        hasStartDate: true,
        currentDay: 3,
        daysSinceParticipation: 2,
        participationDuringTest: 0,
        hasParticipation: false
      },
      expected: TESTER_STATUSES.NEEDS_REMINDER
    },
    {
      facts: {
        currentStatus: TESTER_STATUSES.ACTIVE,
        hasStartDate: true,
        currentDay: 4,
        daysSinceParticipation: 3,
        participationDuringTest: 0,
        hasParticipation: false
      },
      expected: TESTER_STATUSES.AT_RISK
    },
    {
      facts: {
        currentStatus: TESTER_STATUSES.ACTIVE,
        hasStartDate: true,
        currentDay: POST_TEST_DAY,
        daysSinceParticipation: 1,
        participationDuringTest: 1,
        hasParticipation: true
      },
      expected: TESTER_STATUSES.COMPLETED
    },
    {
      facts: {
        currentStatus: TESTER_STATUSES.NEEDS_REMINDER,
        hasStartDate: true,
        currentDay: 2,
        daysSinceParticipation: 1,
        participationDuringTest: 1,
        hasParticipation: true
      },
      expected: TESTER_STATUSES.ACTIVE
    }
  ];

  assertHardeningTest_(
    calculateTestingDayForDates_(start, new Date(Date.UTC(2025, 11, 31, 12))) === 0,
    'A pre-start date must return Day 0.'
  );
  assertHardeningTest_(
    calculateTestingDayForDates_(start, start) === 1,
    'The start date must return Day 1.'
  );
  assertHardeningTest_(
    calculateTestingDayForDates_(start, new Date(Date.UTC(2026, 0, 14, 12))) === 14,
    'The fourteenth calendar date must return Day 14.'
  );
  assertHardeningTest_(
    calculateTestingDayForDates_(start, new Date(Date.UTC(2026, 0, 15, 12))) === POST_TEST_DAY,
    'A post-plan date must return the post-test day.'
  );

  statusScenarios.forEach(function (scenario) {
    assertHardeningTest_(
      recommendTesterStatusFromFacts_(scenario.facts).status === scenario.expected,
      'Status recommendation did not return ' + scenario.expected + '.'
    );
  });

  var payload = validateFeedbackPayload_({
    testerToken: 'test-token',
    activityId: 'test-activity',
    rating: 5,
    completed: 'Yes',
    bugReported: 'No',
    comment: 'Deterministic validation.'
  });
  assertHardeningTest_(
    payload.rating === 5 && payload.completed === true && payload.bugReported === false,
    'Valid feedback payload normalization failed.'
  );
  var invalidFeedbackRejected = false;
  try {
    validateFeedbackPayload_({
      testerToken: 'test-token',
      activityId: 'test-activity',
      rating: 6,
      completed: true,
      bugReported: false,
      comment: ''
    });
  } catch (error) {
    invalidFeedbackRejected = true;
  }
  assertHardeningTest_(invalidFeedbackRejected, 'Invalid feedback rating was accepted.');

  assertHardeningTest_(
    normalizeWhatsAppPhone('+234 801 234 5678') === '2348012345678',
    'WhatsApp phone normalization failed.'
  );
  var localPhoneRejected = false;
  try {
    normalizeWhatsAppPhone('08012345678');
  } catch (error) {
    localPhoneRejected = true;
  }
  assertHardeningTest_(localPhoneRejected, 'A local-format WhatsApp phone was accepted.');
  assertHardeningTest_(
    normalizeWhatsAppEnabled_('TRUE') && !normalizeWhatsAppEnabled_('false'),
    'WhatsApp Enabled normalization failed.'
  );
  assertHardeningTest_(
    resolveWhatsAppRecipientForMode_(
      '447700900001',
      true,
      '2348012345678'
    ) === '2348012345678',
    'WhatsApp Test Mode did not redirect to the test recipient.'
  );
  var missingTestRecipientRejected = false;
  try {
    resolveWhatsAppRecipientForMode_('447700900001', true, '');
  } catch (error) {
    missingTestRecipientRejected = true;
  }
  assertHardeningTest_(
    missingTestRecipientRejected,
    'WhatsApp Test Mode fell back when its test recipient was missing.'
  );
  assertHardeningTest_(
    evolutionResponseConfirmsSend_({ key: { id: 'message-id' } }) &&
      !evolutionResponseConfirmsSend_({}) &&
      evolutionResponseIndicatesFailure_({ success: false }),
    'Evolution response confirmation checks failed.'
  );
  assertHardeningTest_(
    isAlreadyWhatsAppReminded_({ whatsappReminderCount: 1 }) &&
      !isAlreadyWhatsAppReminded_({ whatsappReminderCount: 0 }),
    'WhatsApp duplicate protection checks failed.'
  );

  var currentDate = new Date(Date.UTC(2026, 0, 10, 12));
  var aggregate = calculateDashboardAggregateMetrics_(
    [
      {
        assignedAt: currentDate,
        completedAt: currentDate,
        status: ACTIVITY_STATUSES.COMPLETED
      },
      { assignedAt: currentDate, completedAt: '', status: ACTIVITY_STATUSES.ASSIGNED },
      {
        assignedAt: new Date(Date.UTC(2026, 0, 9, 12)),
        completedAt: new Date(Date.UTC(2026, 0, 9, 12)),
        status: ACTIVITY_STATUSES.COMPLETED
      }
    ],
    [
      { submittedAt: currentDate, rating: 5, bugReported: false },
      {
        submittedAt: new Date(Date.UTC(2026, 0, 9, 12)),
        rating: 3,
        bugReported: true
      },
      { submittedAt: '', rating: 9, bugReported: false }
    ],
    currentDate
  );
  assertHardeningTest_(
    aggregate.today.assignedActivities === 2 &&
      aggregate.today.completedActivities === 1 &&
      aggregate.today.pendingActivities === 1 &&
      aggregate.today.feedbackSubmissions === 1,
    'Deterministic dashboard today metrics are incorrect.'
  );
  assertHardeningTest_(
    aggregate.overall.completedActivities === 2 &&
      dashboardNumbersEqual_(aggregate.overall.completionRate, 2 / 3) &&
      aggregate.overall.feedbackSubmissions === 3 &&
      aggregate.overall.bugsReported === 1 &&
      dashboardNumbersEqual_(aggregate.overall.averageFeedbackRating, 4),
    'Deterministic dashboard overall metrics are incorrect.'
  );

  return {
    ok: true,
    dateScenarios: 4,
    statusScenarios: statusScenarios.length,
    feedbackScenarios: 2,
    whatsAppScenarios: 7,
    dashboardScenarios: 2
  };
}

function runFullSystemDryRun() {
  var configuration = getSystemConfigurationHealth();
  if (!configuration.emailTestMode) {
    throw new Error('Full system dry run requires Email Test Mode to be enabled.');
  }
  if (!configuration.testRecipientValid) {
    throw new Error('Full system dry run requires a valid test email recipient.');
  }
  if (!configuration.feedbackWebAppUrlValid) {
    throw new Error('Full system dry run requires a valid Feedback Web App URL.');
  }

  initializeSpreadsheet();
  var seed = seedDefaultTasks();
  var deterministic = runDeterministicLogicTests();
  var task = getTaskForDayAndPeriod(1, TASK_PERIODS.AM);
  if (!task || !task.active) {
    throw new Error('Full system dry run requires an active Day 1 AM task.');
  }
  var smokeId = Utilities.getUuid().replace(/-/g, '');
  var tester = addTester(
    '[Smoke Test] Full System Dry Run',
    'full-dry-run-' + smokeId + '@example.invalid',
    ''
  );
  tester = activateTester(tester.testerId, new Date());

  var activity = assignTaskToTester(tester.testerId, task.taskId);
  var reminder = sendTaskReminderEmail(tester, task, activity);
  assertHardeningTest_(
    reminder.sent && reminder.isTestMode,
    'Dry-run reminder was not sent in Email Test Mode.'
  );
  var duplicateReminder = sendTaskReminderEmail(tester, task, activity);
  assertHardeningTest_(
    !duplicateReminder.sent && duplicateReminder.reason === 'already_reminded',
    'Dry-run duplicate reminder was not blocked.'
  );

  var feedbackUrl = getFeedbackUrlForActivity(activity.activityId);
  assertHardeningTest_(
    feedbackUrl.indexOf('?tester=') !== -1 && feedbackUrl.indexOf('&activity=') !== -1,
    'Dry-run personalized feedback URL was not generated.'
  );
  var feedback = submitFeedback({
    testerToken: tester.token,
    activityId: activity.activityId,
    rating: 5,
    completed: true,
    bugReported: false,
    comment: '[Smoke Test] Full system dry-run feedback.'
  });
  var duplicateFeedback = submitFeedback({
    testerToken: tester.token,
    activityId: activity.activityId,
    rating: 5,
    completed: true,
    bugReported: false,
    comment: '[Smoke Test] Duplicate full system dry-run feedback.'
  });
  var completedActivity = getActivityById(activity.activityId);
  assertHardeningTest_(
    feedback.ok &&
      !feedback.alreadySubmitted &&
      duplicateFeedback.alreadySubmitted &&
      completedActivity.status === ACTIVITY_STATUSES.COMPLETED,
    'Dry-run feedback lifecycle failed.'
  );

  var status = refreshTesterStatus(tester.testerId);
  var monitoring = refreshMonitoringSheet();
  var dashboard = refreshDashboardFromCurrentMonitoring_();
  assertHardeningTest_(
    monitoringSheetHasTester_(tester.testerId),
    'Dry-run tester was not rendered in Monitoring.'
  );

  return {
    ok: true,
    testerId: tester.testerId,
    taskSeedCreated: seed.created,
    reminderSentInTestMode: reminder.isTestMode,
    duplicateReminderBlocked: duplicateReminder.reason === 'already_reminded',
    feedbackRecorded: !feedback.alreadySubmitted,
    duplicateFeedbackBlocked: duplicateFeedback.alreadySubmitted,
    activityCompleted: completedActivity.status === ACTIVITY_STATUSES.COMPLETED,
    testerStatus: status.status,
    monitoringRowsWritten: monitoring.rowsWritten,
    dashboardTesterCount: dashboard.totalTesters,
    deterministicChecksPassed: deterministic.ok,
    smokeDataRetained: true
  };
}

function createDailyAutomationTrigger_(handler, hour) {
  return ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
}

function getAutomationTriggerSpecs_() {
  return [
    { handler: REMINDER_TRIGGER_HANDLERS.MORNING, hour: 9 },
    { handler: REMINDER_TRIGGER_HANDLERS.EVENING, hour: 18 },
    { handler: MONITORING_TRIGGER_HANDLER, hour: 19 }
  ];
}

function isValidClockTrigger_(trigger) {
  return trigger.getEventType() === ScriptApp.EventType.CLOCK;
}

function inspectRequiredSheetStructure_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];
  var validSheets = {};

  Object.keys(REQUIRED_SHEET_HEADERS).forEach(function (sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    var expectedHeaders = REQUIRED_SHEET_HEADERS[sheetName];
    if (!sheet) {
      results.push({
        sheetName: sheetName,
        ok: false,
        message: sheetName + ' is missing.'
      });
      return;
    }
    if (expectedHeaders.length === 0) {
      validSheets[sheetName] = true;
      results.push({
        sheetName: sheetName,
        ok: true,
        message: sheetName + ' exists.'
      });
      return;
    }

    var useHeaderLookup =
      sheetName === TESTER_SHEET_NAME || sheetName === ACTIVITY_LOG_SHEET_NAME;
    var headerWidth = useHeaderLookup ? sheet.getLastColumn() : expectedHeaders.length;
    var actualHeaders = headerWidth === 0
      ? []
      : sheet
          .getRange(1, 1, 1, headerWidth)
          .getValues()[0]
          .map(function (header) { return String(header).trim(); });
    var mismatches = expectedHeaders.filter(function (header, index) {
      if (!useHeaderLookup) {
        return actualHeaders[index] !== header;
      }
      return actualHeaders.filter(function (candidate) {
        return candidate === header;
      }).length !== 1;
    });
    validSheets[sheetName] = mismatches.length === 0;
    results.push({
      sheetName: sheetName,
      ok: mismatches.length === 0,
      message: mismatches.length === 0
        ? sheetName + ' headers are valid.'
        : sheetName + ' has invalid or missing required headers.'
    });
  });

  return {
    results: results,
    validSheets: validSheets
  };
}

function inspectDefaultTaskCoverage_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SHEET_NAME);
  var headers = REQUIRED_SHEET_HEADERS[TASK_SHEET_NAME];
  var lastRow = sheet.getLastRow();
  var rows = lastRow <= 1
    ? []
    : sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var dayIndex = headers.indexOf('Day');
  var periodIndex = headers.indexOf('Period');
  var activeIndex = headers.indexOf('Active');
  var counts = {};

  rows.forEach(function (row) {
    var day = Number(row[dayIndex]);
    var period = String(row[periodIndex] || '').trim().toUpperCase();
    if (
      day >= 1 &&
      day <= TESTING_PLAN_DAYS &&
      (period === TASK_PERIODS.AM || period === TASK_PERIODS.PM) &&
      isTaskActive_(row[activeIndex])
    ) {
      var key = getTaskDayPeriodKey_(day, period);
      counts[key] = (counts[key] || 0) + 1;
    }
  });

  var missing = [];
  var duplicates = [];
  DEFAULT_TASKS.forEach(function (task) {
    var key = getTaskDayPeriodKey_(task.day, task.period);
    if (!counts[key]) {
      missing.push('Day ' + task.day + ' ' + task.period);
    } else if (counts[key] > 1) {
      duplicates.push('Day ' + task.day + ' ' + task.period);
    }
  });

  return [
    {
      name: 'Default task coverage',
      status: missing.length === 0 ? 'pass' : 'fail',
      message: missing.length === 0
        ? 'All 28 active Day/Period tasks are present.'
        : 'Missing active tasks: ' + missing.join(', ')
    },
    {
      name: 'Default task uniqueness',
      status: duplicates.length === 0 ? 'pass' : 'fail',
      message: duplicates.length === 0
        ? 'No active Day/Period task duplicates were found.'
        : 'Duplicate active tasks: ' + duplicates.join(', ')
    }
  ];
}

function inspectWhatsAppReadiness_() {
  var context = getTesterSheetContext_();
  var lastRow = context.sheet.getLastRow();
  var rows = lastRow <= 1
    ? []
    : context.sheet.getRange(2, 1, lastRow - 1, context.headers.length).getValues();
  var enabledCount = 0;
  var dataIssues = [];
  rows.forEach(function (row, index) {
    if (!normalizeLookupValue_(row[context.columnIndexes['Tester ID'] - 1])) {
      return;
    }
    var enabled;
    try {
      enabled = normalizeWhatsAppEnabled_(
        row[context.columnIndexes['WhatsApp Enabled'] - 1]
      );
    } catch (error) {
      dataIssues.push('Testers row ' + (index + 2) + ' has an invalid WhatsApp Enabled value.');
      return;
    }
    if (!enabled) {
      return;
    }
    enabledCount += 1;
    try {
      if (!normalizeWhatsAppPhone(row[context.columnIndexes['Phone'] - 1])) {
        throw new Error('missing');
      }
    } catch (error) {
      dataIssues.push('Testers row ' + (index + 2) + ' has WhatsApp enabled without a valid phone.');
    }
  });
  var configuration = getWhatsAppConfigurationStatus_();
  var required = [
    { configured: configuration.apiUrlConfigured, label: 'Evolution API URL' },
    { configured: configuration.apiKeyConfigured, label: 'Evolution API key' },
    { configured: configuration.instanceConfigured, label: 'Evolution instance' }
  ];
  var results = [];

  if (dataIssues.length === 0) {
    results.push({
      name: 'WhatsApp tester data',
      status: 'pass',
      message: 'WhatsApp tester phone and opt-in values are valid.'
    });
  } else {
    dataIssues.forEach(function (message) {
      results.push({
        name: 'WhatsApp tester data',
        status: 'fail',
        message: message
      });
    });
  }

  required.forEach(function (item) {
    if (item.configured) {
      results.push({
        name: 'WhatsApp: ' + item.label,
        status: 'pass',
        message: item.label + ' is configured.'
      });
      return;
    }
    results.push({
      name: 'WhatsApp: ' + item.label,
      status: enabledCount > 0 ? 'fail' : 'warning',
      message: item.label + ' is not configured' +
        (enabledCount > 0
          ? ' but at least one tester has WhatsApp enabled.'
          : '; no tester currently has WhatsApp enabled.')
    });
  });

  if (configuration.testMode) {
    results.push({
      name: 'WhatsApp Test Mode recipient',
      status: configuration.testRecipientConfigured && configuration.testRecipientValid
        ? 'pass'
        : 'fail',
      message: configuration.testRecipientConfigured && configuration.testRecipientValid
        ? 'WhatsApp Test Mode has a valid test recipient.'
        : 'WhatsApp Test Mode is enabled but no valid test recipient is configured.'
    });
  } else {
    results.push({
      name: 'WhatsApp Test Mode',
      status: 'pass',
      message: 'WhatsApp Test Mode is OFF.'
    });
  }
  return results;
}

function addReadinessResult_(report, name, status, message) {
  var check = {
    name: name,
    status: status,
    message: message
  };
  report.checks.push(check);
  if (status === 'fail') {
    report.failures.push(message);
  } else if (status === 'warning') {
    report.warnings.push(message);
  }
}

function validateCleanupSheetHeaders_(sheet, sheetName) {
  if (!sheet) {
    throw new Error('The ' + sheetName + ' sheet is missing. Nothing was deleted.');
  }
  var requiredHeaders = REQUIRED_SHEET_HEADERS[sheetName];
  var useHeaderLookup =
    sheetName === TESTER_SHEET_NAME || sheetName === ACTIVITY_LOG_SHEET_NAME;
  var width = useHeaderLookup ? sheet.getLastColumn() : requiredHeaders.length;
  var actual = width === 0
    ? []
    : sheet.getRange(1, 1, 1, width).getValues()[0].map(function (header) {
        return String(header).trim();
      });
  requiredHeaders.forEach(function (header, index) {
    var valid = useHeaderLookup
      ? actual.filter(function (candidate) { return candidate === header; }).length === 1
      : actual[index] === header;
    if (!valid) {
      throw new Error('The ' + sheetName + ' headers are invalid. Nothing was deleted.');
    }
  });
  return useHeaderLookup ? actual : requiredHeaders;
}

function deleteCleanupRowsByTesterId_(sheet, sheetName, eligibleIds) {
  var headers = validateCleanupSheetHeaders_(sheet, sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return 0;
  }
  var testerIdIndex = headers.indexOf('Tester ID');
  var rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var deleted = 0;

  for (var rowIndex = rows.length - 1; rowIndex >= 0; rowIndex -= 1) {
    var testerId = String(rows[rowIndex][testerIdIndex] || '').trim();
    if (Object.prototype.hasOwnProperty.call(eligibleIds, testerId)) {
      sheet.deleteRow(rowIndex + 2);
      deleted += 1;
    }
  }
  return deleted;
}

function isSmokeTestName_(name) {
  return String(name || '').indexOf('[Smoke Test]') === 0;
}

function assertHardeningTest_(condition, message) {
  if (!condition) {
    throw new Error('System hardening test failed: ' + message);
  }
}
