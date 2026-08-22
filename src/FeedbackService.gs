function setFeedbackWebAppUrl(url) {
  if (typeof url !== 'string' || !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(url.trim())) {
    throw new Error('Provide a valid Apps Script Web App deployment URL ending in /exec.');
  }

  PropertiesService.getScriptProperties().setProperty(
    FEEDBACK_WEB_APP_URL_PROPERTY,
    url.trim()
  );
}

function getFeedbackWebAppUrl() {
  var url = PropertiesService.getScriptProperties().getProperty(
    FEEDBACK_WEB_APP_URL_PROPERTY
  );
  if (!url) {
    throw new Error(
      'Feedback Web App URL is not configured. Deploy the Web App, then run setFeedbackWebAppUrl("DEPLOYMENT_URL").'
    );
  }

  return url;
}

function getFeedbackUrlForActivity(activityId) {
  var activity = getActivityById(activityId);
  if (!activity) {
    throw new Error('Activity not found for the provided Activity ID.');
  }

  var tester = getTesterById(activity.testerId);
  if (!tester || !tester.token) {
    throw new Error('The activity does not have a valid tester.');
  }

  validateFeedbackContext_(tester.token, activity.activityId);

  return (
    getFeedbackWebAppUrl() +
    '?tester=' + encodeURIComponent(tester.token) +
    '&activity=' + encodeURIComponent(activity.activityId)
  );
}

function doGet(e) {
  try {
    var parameters = e && e.parameter ? e.parameter : {};
    var token = normalizeFeedbackText_(parameters.tester);
    var activityId = normalizeFeedbackText_(parameters.activity);

    if (!token || !activityId) {
      return renderFeedbackPage_({
        state: 'error',
        message: 'This feedback link is incomplete. Please use the full link from your reminder email.'
      });
    }

    var context = validateFeedbackContext_(token, activityId);
    var existing = getFeedbackForTesterAndTask_(
      context.tester.testerId,
      context.task.taskId
    );
    if (existing) {
      return renderFeedbackPage_({
        state: 'already_submitted',
        message: 'Feedback already submitted. Thank you for helping test Lodge Manager.'
      });
    }

    return renderFeedbackPage_({
      state: 'form',
      testerName: context.tester.name,
      testerToken: token,
      activityId: context.activity.activityId,
      day: context.task.day,
      periodLabel: context.task.period === TASK_PERIODS.AM ? 'Morning' : 'Evening',
      taskTitle: context.task.title,
      taskInstructions: context.task.instructions
    });
  } catch (error) {
    return renderFeedbackPage_({
      state: 'error',
      message: feedbackPublicErrorMessage_(error)
    });
  }
}

function submitFeedback(payload) {
  var normalized = validateFeedbackPayload_(payload);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var context = validateFeedbackContext_(
      normalized.testerToken,
      normalized.activityId
    );
    var existing = getFeedbackForTesterAndTask_(
      context.tester.testerId,
      context.task.taskId
    );
    if (existing) {
      return {
        ok: true,
        alreadySubmitted: true,
        day: context.task.day,
        activityCompleted: context.activity.status === ACTIVITY_STATUSES.COMPLETED
      };
    }

    var feedbackId = Utilities.getUuid();
    var submittedAt = new Date();
    var feedbackContext = getFeedbackSheetContext_();
    var row = buildFeedbackRow_(feedbackContext.headers, {
      'Feedback ID': feedbackId,
      'Tester ID': context.tester.testerId,
      'Task ID': context.task.taskId,
      'Rating': normalized.rating,
      'Completed': normalized.completed,
      'Bug Reported': normalized.bugReported,
      'Comment': normalized.comment,
      'Submitted At': submittedAt
    });

    feedbackContext.sheet
      .getRange(
        feedbackContext.sheet.getLastRow() + 1,
        1,
        1,
        feedbackContext.headers.length
      )
      .setValues([row]);

    var completedActivity = context.activity;
    if (normalized.completed) {
      completedActivity = completeActivityWithoutLock_(context.activity.activityId);
    }

    return {
      ok: true,
      alreadySubmitted: false,
      feedbackId: feedbackId,
      day: context.task.day,
      activityCompleted: completedActivity.status === ACTIVITY_STATUSES.COMPLETED
    };
  } finally {
    lock.releaseLock();
  }
}

function runFeedbackSmokeTest() {
  getFeedbackWebAppUrl();
  seedDefaultTasks();

  var smokeId = Utilities.getUuid().replace(/-/g, '');
  var tester = addTester(
    '[Smoke Test] Feedback Web App Tester',
    'feedback-smoke-' + smokeId + '@example.invalid',
    ''
  );
  tester = activateTester(tester.testerId, new Date());

  var task = getTaskForDayAndPeriod(1, TASK_PERIODS.AM);
  if (!task) {
    throw new Error('Feedback smoke test requires an active Day 1 AM task.');
  }

  var activity = assignTaskToTester(tester.testerId, task.taskId);
  var url = getFeedbackUrlForActivity(activity.activityId);
  var resolvedTester = getTesterByToken(tester.token);
  assertFeedbackSmokeTest_(
    resolvedTester && resolvedTester.testerId === tester.testerId,
    'Token lookup did not resolve the smoke-test tester.'
  );

  var context = validateFeedbackContext_(tester.token, activity.activityId);
  assertFeedbackSmokeTest_(
    context.activity.testerId === tester.testerId,
    'Activity ownership validation failed.'
  );

  var result = submitFeedback({
    testerToken: tester.token,
    activityId: activity.activityId,
    completed: true,
    rating: 5,
    bugReported: false,
    comment: '[Smoke Test] Feedback submission validation.'
  });
  assertFeedbackSmokeTest_(!result.alreadySubmitted, 'First feedback was not recorded.');

  var completedActivity = getActivityById(activity.activityId);
  assertFeedbackSmokeTest_(
    completedActivity.status === ACTIVITY_STATUSES.COMPLETED,
    'Completed feedback did not complete the activity.'
  );

  var duplicate = submitFeedback({
    testerToken: tester.token,
    activityId: activity.activityId,
    completed: true,
    rating: 5,
    bugReported: false,
    comment: '[Smoke Test] Duplicate feedback submission.'
  });
  assertFeedbackSmokeTest_(duplicate.alreadySubmitted, 'Duplicate feedback was not blocked.');

  var previewTask = getTaskForDayAndPeriod(1, TASK_PERIODS.PM);
  if (!previewTask) {
    throw new Error('Feedback smoke test requires an active Day 1 PM task.');
  }
  var previewActivity = assignTaskToTester(tester.testerId, previewTask.taskId);
  var testFormUrl = getFeedbackUrlForActivity(previewActivity.activityId);

  return {
    ok: true,
    testerId: tester.testerId,
    activityId: activity.activityId,
    feedbackId: result.feedbackId,
    personalizedUrlGenerated: url.indexOf('?tester=') !== -1,
    testFormUrl: testFormUrl,
    duplicateBlocked: duplicate.alreadySubmitted,
    activityCompleted: completedActivity.status === ACTIVITY_STATUSES.COMPLETED,
    message: 'Smoke-test tester, activity, and feedback rows remain for manual review.'
  };
}

function validateFeedbackContext_(testerToken, activityId) {
  var tester = getTesterByToken(testerToken);
  if (!tester) {
    throw new Error('INVALID_TESTER');
  }

  var activity = getActivityById(activityId);
  if (!activity) {
    throw new Error('INVALID_ACTIVITY');
  }
  if (activity.testerId !== tester.testerId) {
    throw new Error('ACTIVITY_OWNERSHIP_MISMATCH');
  }
  if (
    activity.status !== ACTIVITY_STATUSES.ASSIGNED &&
    activity.status !== ACTIVITY_STATUSES.COMPLETED &&
    activity.status !== ACTIVITY_STATUSES.SKIPPED
  ) {
    throw new Error('INVALID_ACTIVITY');
  }

  var task = getTaskById(activity.taskId);
  if (!task || task.taskId !== activity.taskId) {
    throw new Error('INVALID_TASK');
  }

  return {
    tester: tester,
    activity: activity,
    task: task
  };
}

function validateFeedbackPayload_(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    throw new Error('Invalid feedback submission.');
  }

  var testerToken = normalizeFeedbackText_(payload.testerToken);
  var activityId = normalizeFeedbackText_(payload.activityId);
  if (!testerToken || !activityId) {
    throw new Error('The feedback link identifiers are missing.');
  }

  var rating = Number(payload.rating);
  if (!isFinite(rating) || Math.floor(rating) !== rating || rating < 1 || rating > 5) {
    throw new Error('Rating must be a whole number from 1 to 5.');
  }

  var completed = normalizeFeedbackBoolean_(payload.completed, 'Completed');
  var bugReported = normalizeFeedbackBoolean_(payload.bugReported, 'Bug Reported');
  var comment = normalizeFeedbackText_(payload.comment);
  if (comment.length > FEEDBACK_COMMENT_MAX_LENGTH) {
    throw new Error(
      'Feedback comment must be ' + FEEDBACK_COMMENT_MAX_LENGTH + ' characters or fewer.'
    );
  }

  return {
    testerToken: testerToken,
    activityId: activityId,
    rating: rating,
    completed: completed,
    bugReported: bugReported,
    comment: comment
  };
}

function getFeedbackSheetContext_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEEDBACK_SHEET_NAME);
  var headers = REQUIRED_SHEET_HEADERS[FEEDBACK_SHEET_NAME];
  if (!sheet) {
    throw new Error('The Feedback sheet is missing. Run initializeSpreadsheet() first.');
  }

  var sheetHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (header, index) {
    if (sheetHeaders[index] !== header) {
      throw new Error('The Feedback sheet headers are invalid. Restore the required headers.');
    }
  });

  var columnIndexes = {};
  headers.forEach(function (header, index) {
    columnIndexes[header] = index + 1;
  });

  return {
    sheet: sheet,
    headers: headers,
    columnIndexes: columnIndexes
  };
}

function getFeedbackForTesterAndTask_(testerId, taskId) {
  var context = getFeedbackSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return null;
  }

  var rows = context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues();
  for (var index = 0; index < rows.length; index += 1) {
    if (
      String(rows[index][context.columnIndexes['Tester ID'] - 1]).trim() === testerId &&
      String(rows[index][context.columnIndexes['Task ID'] - 1]).trim() === taskId
    ) {
      return feedbackObjectFromRow_(rows[index], context.columnIndexes);
    }
  }

  return null;
}

function feedbackObjectFromRow_(row, columnIndexes) {
  return {
    feedbackId: row[columnIndexes['Feedback ID'] - 1],
    testerId: row[columnIndexes['Tester ID'] - 1],
    taskId: row[columnIndexes['Task ID'] - 1],
    rating: row[columnIndexes['Rating'] - 1],
    completed: row[columnIndexes['Completed'] - 1],
    bugReported: row[columnIndexes['Bug Reported'] - 1],
    comment: row[columnIndexes['Comment'] - 1],
    submittedAt: row[columnIndexes['Submitted At'] - 1]
  };
}

function buildFeedbackRow_(headers, valuesByHeader) {
  return headers.map(function (header) {
    return valuesByHeader[header];
  });
}

function renderFeedbackPage_(data) {
  var template = HtmlService.createTemplateFromFile('src/FeedbackPage');
  template.state = data.state || 'error';
  template.message = data.message || '';
  template.testerName = data.testerName || '';
  template.testerToken = data.testerToken || '';
  template.activityId = data.activityId || '';
  template.day = data.day || '';
  template.periodLabel = data.periodLabel || '';
  template.taskTitle = data.taskTitle || '';
  template.taskInstructions = data.taskInstructions || '';

  return template
    .evaluate()
    .setTitle('Lodge Manager Beta Test')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function feedbackPublicErrorMessage_(error) {
  var code = error && error.message ? error.message : '';
  if (code === 'INVALID_TESTER') {
    return 'This feedback link is not valid. Please use the latest link from your reminder email.';
  }
  if (code === 'INVALID_ACTIVITY') {
    return 'This activity could not be found. Please use the latest link from your reminder email.';
  }
  if (code === 'ACTIVITY_OWNERSHIP_MISMATCH') {
    return 'This feedback link does not match the assigned activity.';
  }
  if (code === 'INVALID_TASK') {
    return 'The assigned task is unavailable. Please contact the test coordinator.';
  }

  return 'We could not open this feedback page right now. Please try again later.';
}

function normalizeFeedbackBoolean_(value, fieldName) {
  if (value === true || value === 'true' || value === 'Yes') {
    return true;
  }
  if (value === false || value === 'false' || value === 'No') {
    return false;
  }

  throw new Error(fieldName + ' must be Yes or No.');
}

function normalizeFeedbackText_(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }

  return String(value).trim();
}

function assertFeedbackSmokeTest_(condition, message) {
  if (!condition) {
    throw new Error('Feedback smoke test failed: ' + message);
  }
}
