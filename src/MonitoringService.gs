function getTesterParticipationSummary(testerId) {
  var tester = getTesterById(testerId);
  if (!tester) {
    throw new Error('Tester not found for the provided tester ID.');
  }

  var activities = getActivitiesForTester(tester.testerId);
  var feedbackRecords = getFeedbackRecordsForTester_(tester.testerId);
  var completedDates = activities
    .filter(function (activity) {
      return (
        activity.status === ACTIVITY_STATUSES.COMPLETED &&
        isValidMonitoringDate_(activity.completedAt)
      );
    })
    .map(function (activity) {
      return normalizeMonitoringDate_(activity.completedAt);
    });
  var feedbackDates = feedbackRecords
    .filter(function (feedback) {
      return isValidMonitoringDate_(feedback.submittedAt);
    })
    .map(function (feedback) {
      return normalizeMonitoringDate_(feedback.submittedAt);
    });
  var participationDates = completedDates.concat(feedbackDates);
  var lastParticipationAt = latestMonitoringDate_(participationDates);
  var currentDay = calculateMonitoringCurrentDay_(tester);
  var daysSinceParticipation = calculateDaysSinceParticipation_(
    tester,
    lastParticipationAt
  );

  return {
    testerId: tester.testerId,
    currentDay: currentDay,
    assignedActivities: activities.length,
    completedActivities: completedDates.length,
    feedbackCount: feedbackRecords.length,
    lastParticipationAt: lastParticipationAt,
    daysSinceParticipation: daysSinceParticipation,
    participationDuringTest: countParticipationDuringTest_(
      tester,
      participationDates
    )
  };
}

function getTesterLastActivityDate(testerId) {
  if (!getTesterById(testerId)) {
    throw new Error('Tester not found for the provided tester ID.');
  }

  var completedDates = getActivitiesForTester(testerId)
    .filter(function (activity) {
      return (
        activity.status === ACTIVITY_STATUSES.COMPLETED &&
        isValidMonitoringDate_(activity.completedAt)
      );
    })
    .map(function (activity) {
      return normalizeMonitoringDate_(activity.completedAt);
    });

  return latestMonitoringDate_(completedDates);
}

function getTesterCompletedActivityCount(testerId) {
  if (!getTesterById(testerId)) {
    throw new Error('Tester not found for the provided tester ID.');
  }

  return getActivitiesForTester(testerId).filter(function (activity) {
    return (
      activity.status === ACTIVITY_STATUSES.COMPLETED &&
      isValidMonitoringDate_(activity.completedAt)
    );
  }).length;
}

function getTesterFeedbackCount(testerId) {
  if (!getTesterById(testerId)) {
    throw new Error('Tester not found for the provided tester ID.');
  }

  return getFeedbackRecordsForTester_(testerId).length;
}

function determineTesterStatus(testerId) {
  var tester = getTesterById(testerId);
  if (!tester) {
    throw new Error('Tester not found for the provided tester ID.');
  }

  if (tester.status === TESTER_STATUSES.INACTIVE) {
    return {
      testerId: tester.testerId,
      currentStatus: tester.status,
      recommendedStatus: TESTER_STATUSES.INACTIVE,
      reason: 'Inactive is a manual override and is preserved.'
    };
  }

  var summary = getTesterParticipationSummary(tester.testerId);
  var recommendation = recommendTesterStatusFromFacts_({
    currentStatus: tester.status,
    hasStartDate: isValidMonitoringDate_(tester.startDate),
    currentDay: summary.currentDay,
    daysSinceParticipation: summary.daysSinceParticipation,
    participationDuringTest: summary.participationDuringTest,
    hasParticipation: Boolean(summary.lastParticipationAt)
  });
  return statusRecommendation_(
    tester,
    summary,
    recommendation.status,
    recommendation.reason
  );
}

function recommendTesterStatusFromFacts_(facts) {
  if (facts.currentStatus === TESTER_STATUSES.INACTIVE) {
    return {
      status: TESTER_STATUSES.INACTIVE,
      reason: 'Inactive is a manual override and is preserved.'
    };
  }
  if (!facts.hasStartDate) {
    return {
      status: TESTER_STATUSES.NOT_STARTED,
      reason: 'No Start Date is configured.'
    };
  }
  if (facts.currentDay === 0) {
    return {
      status: TESTER_STATUSES.NOT_STARTED,
      reason: 'The testing Start Date is in the future.'
    };
  }
  if (facts.currentDay > TESTING_PLAN_DAYS) {
    return facts.participationDuringTest > 0
      ? {
          status: TESTER_STATUSES.COMPLETED,
          reason: 'The 14-day period ended with recorded participation.'
        }
      : {
          status: TESTER_STATUSES.AT_RISK,
          reason: 'The 14-day period ended without recorded participation.'
        };
  }
  if (facts.daysSinceParticipation >= 3) {
    return {
      status: TESTER_STATUSES.AT_RISK,
      reason:
        'No completed activity or feedback for ' +
        facts.daysSinceParticipation +
        ' calendar days.'
    };
  }
  if (facts.daysSinceParticipation === 2) {
    return {
      status: TESTER_STATUSES.NEEDS_REMINDER,
      reason: 'No completed activity or feedback for 2 calendar days.'
    };
  }
  return {
    status: TESTER_STATUSES.ACTIVE,
    reason: facts.hasParticipation
      ? 'Meaningful participation occurred within the previous 1 calendar day.'
      : 'The tester is within the initial 1-day participation grace period.'
  };
}

function refreshTesterStatus(testerId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var recommendation = determineTesterStatus(testerId);
    if (recommendation.currentStatus === TESTER_STATUSES.INACTIVE) {
      return {
        testerId: testerId,
        changed: false,
        previousStatus: TESTER_STATUSES.INACTIVE,
        status: TESTER_STATUSES.INACTIVE,
        reason: recommendation.reason
      };
    }

    if (recommendation.currentStatus === recommendation.recommendedStatus) {
      return {
        testerId: testerId,
        changed: false,
        previousStatus: recommendation.currentStatus,
        status: recommendation.currentStatus,
        reason: recommendation.reason
      };
    }

    var updated = updateTesterWithoutLock_(testerId, {
      status: recommendation.recommendedStatus
    });
    return {
      testerId: testerId,
      changed: true,
      previousStatus: recommendation.currentStatus,
      status: updated.status,
      reason: recommendation.reason
    };
  } finally {
    lock.releaseLock();
  }
}

function refreshAllTesterStatuses() {
  var result = {
    processed: 0,
    changed: 0,
    unchanged: 0,
    failed: 0,
    changes: [],
    errors: []
  };

  getAllTesters().forEach(function (tester) {
    result.processed += 1;
    try {
      var refresh = refreshTesterStatus(tester.testerId);
      if (refresh.changed) {
        result.changed += 1;
        result.changes.push(refresh);
      } else {
        result.unchanged += 1;
      }
    } catch (error) {
      result.failed += 1;
      result.errors.push({
        testerId: tester.testerId,
        message: error.message
      });
    }
  });

  result.statusCounts = countTesterStatuses_();
  return result;
}

function refreshMonitoringSheet() {
  var context = getMonitoringSheetContext_();
  var testers = getAllTesters();
  var rows = [];
  var errors = [];
  var now = new Date();

  testers.forEach(function (tester) {
    try {
      var summary = getTesterParticipationSummary(tester.testerId);
      rows.push(
        buildMonitoringRow_(context.headers, {
          'Tester ID': tester.testerId,
          'Name': tester.name,
          'Email': tester.email,
          'Current Day': summary.currentDay,
          'Status': tester.status,
          'Completed Activities': summary.completedActivities,
          'Feedback Count': summary.feedbackCount,
          'Last Participation At': summary.lastParticipationAt || '',
          'Days Since Participation':
            summary.daysSinceParticipation === null
              ? ''
              : summary.daysSinceParticipation,
          'Updated At': now
        })
      );
    } catch (error) {
      errors.push({
        testerId: tester.testerId,
        message: error.message
      });
      rows.push(
        buildMonitoringRow_(context.headers, {
          'Tester ID': tester.testerId,
          'Name': tester.name,
          'Email': tester.email,
          'Current Day': '',
          'Status': tester.status,
          'Completed Activities': '',
          'Feedback Count': '',
          'Last Participation At': '',
          'Days Since Participation': '',
          'Updated At': now
        })
      );
    }
  });

  var lastRow = context.sheet.getLastRow();
  if (lastRow > 1) {
    context.sheet
      .getRange(2, 1, lastRow - 1, context.headers.length)
      .clearContent();
  }
  if (rows.length > 0) {
    context.sheet
      .getRange(2, 1, rows.length, context.headers.length)
      .setValues(rows);
  }

  return {
    processed: testers.length,
    rowsWritten: rows.length,
    failed: errors.length,
    errors: errors,
    statusCounts: countTesterStatuses_()
  };
}

function runMonitoringRefresh() {
  var statuses = refreshAllTesterStatuses();
  var monitoring = refreshMonitoringSheet();
  var dashboard = null;
  var dashboardErrors = [];
  try {
    dashboard = refreshDashboardFromCurrentMonitoring_();
  } catch (error) {
    dashboardErrors.push({
      component: 'Dashboard',
      message: error.message
    });
  }

  return {
    processed: statuses.processed,
    changed: statuses.changed,
    unchanged: statuses.unchanged,
    failed: statuses.failed + monitoring.failed + dashboardErrors.length,
    changes: statuses.changes,
    errors: statuses.errors.concat(monitoring.errors, dashboardErrors),
    rowsWritten: monitoring.rowsWritten,
    statusCounts: monitoring.statusCounts,
    dashboard: dashboard
  };
}

function installMonitoringTrigger() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var existing = ScriptApp.getProjectTriggers().some(function (trigger) {
      return trigger.getHandlerFunction() === MONITORING_TRIGGER_HANDLER;
    });
    if (existing) {
      return {
        created: false,
        existing: true
      };
    }

    createDailyAutomationTrigger_(MONITORING_TRIGGER_HANDLER, 19);

    return {
      created: true,
      existing: false
    };
  } finally {
    lock.releaseLock();
  }
}

function removeMonitoringTrigger() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var removed = 0;
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
      if (trigger.getHandlerFunction() === MONITORING_TRIGGER_HANDLER) {
        ScriptApp.deleteTrigger(trigger);
        removed += 1;
      }
    });

    return {
      removed: removed
    };
  } finally {
    lock.releaseLock();
  }
}

function runMonitoringSmokeTest() {
  initializeSpreadsheet();
  seedDefaultTasks();
  var today = new Date();
  var smokeId = Utilities.getUuid().replace(/-/g, '');
  var dayOneTask = getTaskForDayAndPeriod(1, TASK_PERIODS.AM);
  if (!dayOneTask) {
    throw new Error('Monitoring smoke test requires an active Day 1 AM task.');
  }

  var future = addTester(
    '[Smoke Test] Monitoring Future Start',
    'monitor-future-' + smokeId + '@example.invalid',
    addMonitoringCalendarDays_(today, 1)
  );

  var active = addTester(
    '[Smoke Test] Monitoring Active',
    'monitor-active-' + smokeId + '@example.invalid',
    ''
  );
  active = activateTester(active.testerId, today);
  var activeActivity = assignTaskToTester(active.testerId, dayOneTask.taskId);
  completeActivity(activeActivity.activityId);

  var needsReminder = addTester(
    '[Smoke Test] Monitoring Needs Reminder',
    'monitor-needs-' + smokeId + '@example.invalid',
    ''
  );
  needsReminder = activateTester(
    needsReminder.testerId,
    addMonitoringCalendarDays_(today, -2)
  );

  var atRisk = addTester(
    '[Smoke Test] Monitoring At Risk',
    'monitor-risk-' + smokeId + '@example.invalid',
    ''
  );
  atRisk = activateTester(
    atRisk.testerId,
    addMonitoringCalendarDays_(today, -3)
  );

  var completed = addTester(
    '[Smoke Test] Monitoring Completed',
    'monitor-completed-' + smokeId + '@example.invalid',
    ''
  );
  var completedStart = addMonitoringCalendarDays_(today, -15);
  completed = activateTester(completed.testerId, completedStart);
  var historicalActivity = assignTaskToTester(completed.testerId, dayOneTask.taskId);
  setMonitoringSmokeActivityDate_(
    historicalActivity.activityId,
    addMonitoringCalendarDays_(completedStart, 1)
  );

  var inactive = addTester(
    '[Smoke Test] Monitoring Inactive',
    'monitor-inactive-' + smokeId + '@example.invalid',
    ''
  );
  inactive = updateTester(inactive.testerId, {
    startDate: addMonitoringCalendarDays_(today, -3),
    status: TESTER_STATUSES.INACTIVE
  });

  var scenarios = [
    { tester: future, expected: TESTER_STATUSES.NOT_STARTED },
    { tester: active, expected: TESTER_STATUSES.ACTIVE },
    { tester: needsReminder, expected: TESTER_STATUSES.NEEDS_REMINDER },
    { tester: atRisk, expected: TESTER_STATUSES.AT_RISK },
    { tester: completed, expected: TESTER_STATUSES.COMPLETED },
    { tester: inactive, expected: TESTER_STATUSES.INACTIVE }
  ];

  var refreshResults = scenarios.map(function (scenario) {
    var result = refreshTesterStatus(scenario.tester.testerId);
    assertMonitoringSmokeTest_(
      result.status === scenario.expected,
      scenario.tester.name + ' should be ' + scenario.expected + '.'
    );
    return result;
  });

  var idempotent = refreshTesterStatus(needsReminder.testerId);
  assertMonitoringSmokeTest_(
    !idempotent.changed && idempotent.status === TESTER_STATUSES.NEEDS_REMINDER,
    'Status refresh should be idempotent.'
  );
  var preservedInactive = refreshTesterStatus(inactive.testerId);
  assertMonitoringSmokeTest_(
    !preservedInactive.changed && preservedInactive.status === TESTER_STATUSES.INACTIVE,
    'Inactive manual override was not preserved.'
  );

  var monitoring = refreshMonitoringSheet();
  assertMonitoringSmokeTest_(
    monitoringSheetHasTester_(active.testerId),
    'Monitoring sheet row was not generated.'
  );

  return {
    ok: true,
    scenarioCount: scenarios.length,
    changed: refreshResults.filter(function (result) {
      return result.changed;
    }).length,
    idempotent: !idempotent.changed,
    inactivePreserved: preservedInactive.status === TESTER_STATUSES.INACTIVE,
    monitoringRowGenerated: true,
    statuses: scenarios.map(function (scenario) {
      return {
        testerId: scenario.tester.testerId,
        expected: scenario.expected,
        actual: getTesterById(scenario.tester.testerId).status
      };
    }),
    message: 'Smoke-test rows remain for manual review. Real tester source rows were not changed.'
  };
}

function getFeedbackRecordsForTester_(testerId) {
  var context = getFeedbackSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues()
    .filter(function (row) {
      return String(row[context.columnIndexes['Tester ID'] - 1]).trim() === testerId;
    })
    .map(function (row) {
      return feedbackObjectFromRow_(row, context.columnIndexes);
    });
}

function calculateMonitoringCurrentDay_(tester) {
  if (!isValidMonitoringDate_(tester.startDate)) {
    return 0;
  }
  return calculateTestingDayForDates_(normalizeMonitoringDate_(tester.startDate), new Date());
}

function calculateDaysSinceParticipation_(tester, lastParticipationAt) {
  if (lastParticipationAt) {
    return Math.max(
      0,
      monitoringCalendarDayDifference_(lastParticipationAt, new Date())
    );
  }
  if (!isValidMonitoringDate_(tester.startDate)) {
    return null;
  }

  var sinceStart = monitoringCalendarDayDifference_(tester.startDate, new Date());
  return sinceStart < 0 ? null : sinceStart;
}

function countParticipationDuringTest_(tester, participationDates) {
  if (!isValidMonitoringDate_(tester.startDate)) {
    return 0;
  }

  var start = calendarDateAsUtcMillis_(normalizeMonitoringDate_(tester.startDate));
  var end = start + (TESTING_PLAN_DAYS - 1) * 24 * 60 * 60 * 1000;
  return participationDates.filter(function (date) {
    var calendarDate = calendarDateAsUtcMillis_(date);
    return calendarDate >= start && calendarDate <= end;
  }).length;
}

function statusRecommendation_(tester, summary, status, reason) {
  return {
    testerId: tester.testerId,
    currentStatus: tester.status,
    recommendedStatus: status,
    currentDay: summary.currentDay,
    reason: reason
  };
}

function getMonitoringSheetContext_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MONITORING_SHEET_NAME);
  var headers = REQUIRED_SHEET_HEADERS[MONITORING_SHEET_NAME];
  if (!sheet) {
    throw new Error('The Monitoring sheet is missing. Run initializeSpreadsheet() first.');
  }

  var sheetHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (header, index) {
    if (sheetHeaders[index] !== header) {
      throw new Error('The Monitoring sheet headers are invalid. Restore the required headers.');
    }
  });

  return {
    sheet: sheet,
    headers: headers
  };
}

function buildMonitoringRow_(headers, valuesByHeader) {
  return headers.map(function (header) {
    return valuesByHeader[header];
  });
}

function countTesterStatuses_() {
  var counts = {};
  Object.keys(TESTER_STATUSES).forEach(function (key) {
    counts[TESTER_STATUSES[key]] = 0;
  });
  getAllTesters().forEach(function (tester) {
    if (Object.prototype.hasOwnProperty.call(counts, tester.status)) {
      counts[tester.status] += 1;
    }
  });
  return counts;
}

function monitoringSheetHasTester_(testerId) {
  var context = getMonitoringSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return false;
  }

  return context.sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .some(function (row) {
      return String(row[0]).trim() === testerId;
    });
}

function setMonitoringSmokeActivityDate_(activityId, date) {
  var match = findActivityRowById_(activityId);
  if (!match) {
    throw new Error('Smoke-test activity could not be found.');
  }

  match.sheet
    .getRange(match.rowNumber, match.columnIndexes['Assigned At'])
    .setValue(date);
  match.sheet
    .getRange(match.rowNumber, match.columnIndexes['Completed At'])
    .setValue(date);
  match.sheet
    .getRange(match.rowNumber, match.columnIndexes['Status'])
    .setValue(ACTIVITY_STATUSES.COMPLETED);
}

function addMonitoringCalendarDays_(date, days) {
  var targetCalendar =
    calendarDateAsUtcMillis_(normalizeMonitoringDate_(date)) +
    days * 24 * 60 * 60 * 1000;
  var candidate = new Date(targetCalendar + 12 * 60 * 60 * 1000);
  var candidateCalendar = calendarDateAsUtcMillis_(candidate);
  candidate.setTime(candidate.getTime() + targetCalendar - candidateCalendar);
  return candidate;
}

function monitoringCalendarDayDifference_(earlierDate, laterDate) {
  return Math.round(
    (calendarDateAsUtcMillis_(normalizeMonitoringDate_(laterDate)) -
      calendarDateAsUtcMillis_(normalizeMonitoringDate_(earlierDate))) /
      (24 * 60 * 60 * 1000)
  );
}

function latestMonitoringDate_(dates) {
  if (dates.length === 0) {
    return null;
  }

  return dates.reduce(function (latest, date) {
    return date.getTime() > latest.getTime() ? date : latest;
  });
}

function isValidMonitoringDate_(value) {
  if (value === null || value === '' || typeof value === 'undefined') {
    return false;
  }
  var date = value instanceof Date ? value : new Date(value);
  return !isNaN(date.getTime());
}

function normalizeMonitoringDate_(value) {
  var date = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  if (isNaN(date.getTime())) {
    throw new Error('A valid monitoring date is required.');
  }
  return date;
}

function assertMonitoringSmokeTest_(condition, message) {
  if (!condition) {
    throw new Error('Monitoring smoke test failed: ' + message);
  }
}
