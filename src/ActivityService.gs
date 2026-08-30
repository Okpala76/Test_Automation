function generateActivityId() {
  return Utilities.getUuid();
}

function assignTaskToTester(testerId, taskId) {
  return assignTaskToTesterWithResult_(testerId, taskId).activity;
}

function assignTodayTasksForTester(testerId) {
  var today = getTodayTasksForTester(testerId);
  var summary = {
    testerId: today.testerId,
    day: today.currentDay,
    assigned: 0,
    existing: 0,
    taskCount: today.tasks.length
  };

  today.tasks.forEach(function (task) {
    var assignment = assignTaskToTesterWithResult_(today.testerId, task.taskId);
    if (assignment.created) {
      summary.assigned += 1;
    } else {
      summary.existing += 1;
    }
  });

  return summary;
}

function assignTodayTasksForAllActiveTesters() {
  var result = {
    processed: 0,
    assigned: 0,
    existing: 0,
    noTasks: 0,
    errors: []
  };

  getActiveTesters().forEach(function (tester) {
    result.processed += 1;
    try {
      var summary = assignTodayTasksForTester(tester.testerId);
      result.assigned += summary.assigned;
      result.existing += summary.existing;
      if (summary.taskCount === 0) {
        result.noTasks += 1;
      }
    } catch (error) {
      result.errors.push({
        testerId: tester.testerId,
        message: error.message
      });
    }
  });

  return result;
}

function getActivityById(activityId) {
  var normalizedId = normalizeActivityLookupValue_(activityId);
  if (!normalizedId) {
    return null;
  }

  var match = findActivityRowById_(normalizedId);
  return match ? activityObjectFromRow_(match.row, match.columnIndexes) : null;
}

function getActivitiesForTester(testerId) {
  var normalizedTesterId = normalizeActivityLookupValue_(testerId);
  if (!normalizedTesterId) {
    return [];
  }

  return getAllActivities_().filter(function (activity) {
    return activity.testerId === normalizedTesterId;
  });
}

function getActivityForTesterAndTask(testerId, taskId) {
  var normalizedTesterId = normalizeActivityLookupValue_(testerId);
  var normalizedTaskId = normalizeActivityLookupValue_(taskId);
  if (!normalizedTesterId || !normalizedTaskId) {
    return null;
  }

  var activities = getAllActivities_();
  for (var index = 0; index < activities.length; index += 1) {
    if (
      activities[index].testerId === normalizedTesterId &&
      activities[index].taskId === normalizedTaskId
    ) {
      return activities[index];
    }
  }

  return null;
}

function completeActivity(activityId) {
  var normalizedId = normalizeActivityLookupValue_(activityId);
  if (!normalizedId) {
    throw new Error('An Activity ID is required.');
  }

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    return completeActivityWithoutLock_(normalizedId);
  } finally {
    lock.releaseLock();
  }
}

function completeActivityWithoutLock_(activityId) {
  var match = findActivityRowById_(activityId);
  if (!match) {
    throw new Error('Activity not found for the provided Activity ID.');
  }

  var activity = activityObjectFromRow_(match.row, match.columnIndexes);
  if (activity.status === ACTIVITY_STATUSES.COMPLETED) {
    return activity;
  }

  match.sheet
    .getRange(match.rowNumber, match.columnIndexes['Status'])
    .setValue(ACTIVITY_STATUSES.COMPLETED);
  match.sheet
    .getRange(match.rowNumber, match.columnIndexes['Completed At'])
    .setValue(new Date());

  return getActivityById(activityId);
}

function runTaskEngineSmokeTest() {
  var firstSeed = seedDefaultTasks();
  var secondSeed = seedDefaultTasks();
  assertTaskEngineSmokeTest_(
    DEFAULT_TASKS.length === TESTING_PLAN_DAYS * 2,
    'The default plan must contain 28 tasks.'
  );
  assertTaskEngineSmokeTest_(secondSeed.created === 0, 'Task seeding created duplicates.');
  DEFAULT_TASKS.forEach(function (task) {
    assertTaskEngineSmokeTest_(
      getTaskForDayAndPeriod(task.day, task.period) !== null,
      'A default Day ' + task.day + ' ' + task.period + ' task is missing.'
    );
  });

  var smokeId = Utilities.getUuid().replace(/-/g, '');
  var email = 'task-smoke-' + smokeId + '@example.invalid';
  var tester = addTester('[Smoke Test] Task Engine Tester', email, '');
  var activatedTester = activateTester(tester.testerId, new Date());
  var currentDay = getTesterCurrentDay(activatedTester.testerId);
  assertTaskEngineSmokeTest_(currentDay === 1, 'A newly activated tester must be on Day 1.');

  var today = getTodayTasksForTester(activatedTester.testerId);
  assertTaskEngineSmokeTest_(today.tasks.length > 0, 'No active task was found for Day 1.');

  var firstAssignment = assignTaskToTesterWithResult_(
    activatedTester.testerId,
    today.tasks[0].taskId
  );
  var duplicateAssignment = assignTaskToTesterWithResult_(
    activatedTester.testerId,
    today.tasks[0].taskId
  );
  assertTaskEngineSmokeTest_(firstAssignment.created, 'The first assignment was not created.');
  assertTaskEngineSmokeTest_(
    !duplicateAssignment.created &&
      duplicateAssignment.activity.activityId === firstAssignment.activity.activityId,
    'Duplicate assignment prevention failed.'
  );

  var todayAssignmentSummary = assignTodayTasksForTester(activatedTester.testerId);
  var completedActivity = completeActivity(firstAssignment.activity.activityId);
  assertTaskEngineSmokeTest_(
    completedActivity.status === ACTIVITY_STATUSES.COMPLETED &&
      completedActivity.completedAt instanceof Date,
    'Activity completion failed.'
  );

  return {
    ok: true,
    seedCreated: firstSeed.created,
    seedSkipped: firstSeed.skipped,
    testerId: activatedTester.testerId,
    day: currentDay,
    taskCount: today.tasks.length,
    assigned: todayAssignmentSummary.assigned,
    existing: todayAssignmentSummary.existing,
    completedActivityId: completedActivity.activityId,
    message: 'Smoke-test tester and activity rows remain for manual review.'
  };
}

function assignTaskToTesterWithResult_(testerId, taskId) {
  var normalizedTesterId = normalizeActivityLookupValue_(testerId);
  var normalizedTaskId = normalizeActivityLookupValue_(taskId);
  if (!normalizedTesterId || !normalizedTaskId) {
    throw new Error('Tester ID and Task ID are required for assignment.');
  }

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    return assignTaskToTesterWithResultWithoutLock_(
      normalizedTesterId,
      normalizedTaskId
    );
  } finally {
    lock.releaseLock();
  }
}

function assignTaskToTesterWithResultWithoutLock_(testerId, taskId) {
  var tester = getTesterById(testerId);
  if (!tester) {
    throw new Error('Tester not found for the provided tester ID.');
  }

  var task = getTaskById(taskId);
  if (!task) {
    throw new Error('Task not found for the provided task ID.');
  }
  if (!task.active) {
    throw new Error('Inactive tasks cannot be assigned.');
  }

  var existingActivity = getActivityForTesterAndTask(testerId, taskId);
  if (existingActivity) {
    return {
      activity: existingActivity,
      created: false
    };
  }

  var context = getActivityLogSheetContext_();
  var activityId = generateUniqueActivityId_();
  var row = buildActivityRow_(context.headers, {
    'Activity ID': activityId,
    'Tester ID': testerId,
    'Task ID': taskId,
    'Assigned At': new Date(),
    'Completed At': '',
    'Status': ACTIVITY_STATUSES.ASSIGNED,
    'Last Reminder At': '',
    'Reminder Count': 0,
    'Last WhatsApp Reminder At': '',
    'WhatsApp Reminder Count': 0
  });

  context.sheet
    .getRange(context.sheet.getLastRow() + 1, 1, 1, context.headers.length)
    .setValues([row]);

  return {
    activity: activityObjectFromRow_(row, context.columnIndexes),
    created: true
  };
}

function ensureActivityLogReminderColumns_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVITY_LOG_SHEET_NAME);
  if (!sheet) {
    return;
  }
  var required = REQUIRED_SHEET_HEADERS[ACTIVITY_LOG_SHEET_NAME];
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    return;
  }
  var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (value) {
    return String(value).trim();
  });
  var missing = [];
  required.forEach(function (header) {
    if (currentHeaders.indexOf(header) === -1) {
      missing.push(header);
    }
  });
  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
}

function getActivityLogSheetContext_() {
  ensureActivityLogReminderColumns_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVITY_LOG_SHEET_NAME);
  var requiredHeaders = REQUIRED_SHEET_HEADERS[ACTIVITY_LOG_SHEET_NAME];

  if (!sheet) {
    throw new Error('The Activity Log sheet is missing. Run initializeSpreadsheet() first.');
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    throw new Error('The Activity Log headers are invalid. Run initializeSpreadsheet() first.');
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (header) {
    return String(header).trim();
  });
  requiredHeaders.forEach(function (header) {
    if (headers.filter(function (candidate) { return candidate === header; }).length !== 1) {
      throw new Error(
        'The Activity Log headers are invalid. Run initializeSpreadsheet() or restore the required headers.'
      );
    }
  });

  var columnIndexes = {};
  requiredHeaders.forEach(function (header) {
    columnIndexes[header] = headers.indexOf(header) + 1;
  });

  return {
    sheet: sheet,
    headers: headers,
    columnIndexes: columnIndexes
  };
}

function getAllActivities_() {
  var context = getActivityLogSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues()
    .filter(function (row) {
      return normalizeActivityLookupValue_(row[context.columnIndexes['Activity ID'] - 1]) !== '';
    })
    .map(function (row) {
      return activityObjectFromRow_(row, context.columnIndexes);
    });
}

function findActivityRowById_(activityId) {
  var context = getActivityLogSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return null;
  }

  var values = context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues();
  var activityIdColumn = context.columnIndexes['Activity ID'] - 1;

  for (var index = 0; index < values.length; index += 1) {
    if (normalizeActivityLookupValue_(values[index][activityIdColumn]) === activityId) {
      return {
        sheet: context.sheet,
        headers: context.headers,
        columnIndexes: context.columnIndexes,
        row: values[index],
        rowNumber: index + 2
      };
    }
  }

  return null;
}

function buildActivityRow_(headers, valuesByHeader) {
  return headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(valuesByHeader, header)
      ? valuesByHeader[header]
      : '';
  });
}

function activityObjectFromRow_(row, columnIndexes) {
  return {
    activityId: normalizeActivityLookupValue_(row[columnIndexes['Activity ID'] - 1]),
    testerId: normalizeActivityLookupValue_(row[columnIndexes['Tester ID'] - 1]),
    taskId: normalizeActivityLookupValue_(row[columnIndexes['Task ID'] - 1]),
    assignedAt: row[columnIndexes['Assigned At'] - 1],
    completedAt: row[columnIndexes['Completed At'] - 1],
    status: row[columnIndexes['Status'] - 1],
    lastReminderAt: row[columnIndexes['Last Reminder At'] - 1],
    reminderCount: normalizeReminderCount_(row[columnIndexes['Reminder Count'] - 1]),
    lastWhatsAppReminderAt: row[columnIndexes['Last WhatsApp Reminder At'] - 1],
    whatsappReminderCount: normalizeReminderCount_(
      row[columnIndexes['WhatsApp Reminder Count'] - 1]
    )
  };
}

function normalizeReminderCount_(value) {
  if (value === null || typeof value === 'undefined' || String(value).trim() === '') {
    return 0;
  }
  var parsed = Number(value);
  return isFinite(parsed) && parsed >= 0 ? Math.floor(parsed) : 0;
}

function recordReminderSent_(activityId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    return recordReminderSentWithoutLock_(activityId);
  } finally {
    lock.releaseLock();
  }
}

function recordReminderSentWithoutLock_(activityId) {
  var match = findActivityRowById_(activityId);
  if (!match) {
    throw new Error('Activity not found for the provided Activity ID.');
  }
  var now = new Date();
  var currentCount = normalizeReminderCount_(
    match.row[match.columnIndexes['Reminder Count'] - 1]
  );
  match.sheet
    .getRange(match.rowNumber, match.columnIndexes['Last Reminder At'])
    .setValue(now);
  match.sheet
    .getRange(match.rowNumber, match.columnIndexes['Reminder Count'])
    .setValue(currentCount + 1);
  return getActivityById(activityId);
}

function isAlreadyRemindedToday_(activity) {
  var last = activity.lastReminderAt;
  if (!last) {
    return false;
  }
  var lastDate = last instanceof Date ? last : new Date(last);
  if (!(lastDate instanceof Date) || isNaN(lastDate.getTime())) {
    return false;
  }
  var tz = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var lastStr = Utilities.formatDate(lastDate, tz, 'yyyy-MM-dd');
  return todayStr === lastStr;
}

function recordWhatsAppReminderSent_(activityId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return recordWhatsAppReminderSentWithoutLock_(activityId);
  } finally {
    lock.releaseLock();
  }
}

function recordWhatsAppReminderSentWithoutLock_(activityId) {
  var match = findActivityRowById_(activityId);
  if (!match) {
    throw new Error('Activity not found for the provided Activity ID.');
  }
  var currentCount = normalizeReminderCount_(
    match.row[match.columnIndexes['WhatsApp Reminder Count'] - 1]
  );
  var timestampColumn = match.columnIndexes['Last WhatsApp Reminder At'];
  var countColumn = match.columnIndexes['WhatsApp Reminder Count'];
  if (countColumn === timestampColumn + 1) {
    match.sheet
      .getRange(match.rowNumber, timestampColumn, 1, 2)
      .setValues([[new Date(), currentCount + 1]]);
  } else {
    match.sheet.getRange(match.rowNumber, timestampColumn).setValue(new Date());
    match.sheet.getRange(match.rowNumber, countColumn).setValue(currentCount + 1);
  }
  return getActivityById(activityId);
}

function isAlreadyWhatsAppReminded_(activity) {
  return normalizeReminderCount_(activity && activity.whatsappReminderCount) > 0;
}

function generateUniqueActivityId_() {
  for (var attempt = 0; attempt < 10; attempt += 1) {
    var activityId = generateActivityId();
    if (!getActivityById(activityId)) {
      return activityId;
    }
  }

  throw new Error('Unable to generate a unique activity ID. Please try again.');
}

function normalizeActivityLookupValue_(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }

  return String(value).trim();
}

function assertTaskEngineSmokeTest_(condition, message) {
  if (!condition) {
    throw new Error('Task engine smoke test failed: ' + message);
  }
}
