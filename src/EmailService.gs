function setEmailTestMode(enabled) {
  PropertiesService.getScriptProperties().setProperty(
    EMAIL_TEST_MODE_PROPERTY,
    enabled ? 'true' : 'false'
  );
}

function getEmailTestMode() {
  return PropertiesService.getScriptProperties().getProperty(EMAIL_TEST_MODE_PROPERTY) === 'true';
}

function isEmailTestModeEnabled_() {
  return getEmailTestMode();
}

function setTestEmailRecipient(email) {
  if (typeof email !== 'string' || email.trim() === '') {
    throw new Error('A test email recipient is required.');
  }
  var normalized = email.trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized)) {
    throw new Error('A valid test email recipient is required.');
  }
  PropertiesService.getScriptProperties().setProperty(TEST_EMAIL_RECIPIENT_PROPERTY, normalized);
}

function getTestEmailRecipient() {
  return PropertiesService.getScriptProperties().getProperty(TEST_EMAIL_RECIPIENT_PROPERTY) || '';
}

function resolveEmailRecipient_(testerEmail) {
  if (isEmailTestModeEnabled_()) {
    var testRecipient = getTestEmailRecipient();
    if (!testRecipient) {
      throw new Error('Test email recipient is not configured. Run setTestEmailRecipient("you@example.com") first.');
    }
    return testRecipient;
  }
  return testerEmail;
}

function sendTaskReminderEmail(tester, task, activity) {
  if (!tester || !tester.email) {
    throw new Error('A valid tester with an email is required.');
  }
  if (!task || !task.taskId || !task.title) {
    throw new Error('A valid task is required.');
  }
  if (!activity || !activity.activityId) {
    throw new Error('A valid activity is required.');
  }
  return runWithReminderDeliveryLocks_(function () {
    var currentTester = getTesterById(tester.testerId);
    if (
      !currentTester ||
      REMINDER_ELIGIBLE_STATUSES.indexOf(currentTester.status) === -1
    ) {
      return {
        sent: false,
        reason: 'not_eligible',
        activityId: activity.activityId
      };
    }
    var currentTask = getTaskById(task.taskId);
    if (!currentTask || !currentTask.active) {
      return {
        sent: false,
        reason: 'no_task',
        activityId: activity.activityId
      };
    }
    tester = currentTester;
    task = currentTask;
    activity = getActivityById(activity.activityId);
    if (
      !activity ||
      activity.testerId !== tester.testerId ||
      activity.taskId !== task.taskId
    ) {
      throw new Error('The reminder activity does not match the tester and task.');
    }
    if (activity.status === ACTIVITY_STATUSES.COMPLETED) {
      return {
        sent: false,
        reason: 'already_completed',
        activityId: activity.activityId
      };
    }
    if (isAlreadyRemindedToday_(activity)) {
      return {
        sent: false,
        reason: 'already_reminded',
        activityId: activity.activityId
      };
    }

    var recipient = resolveEmailRecipient_(tester.email);
    var isTestMode = isEmailTestModeEnabled_();
    var subject = buildReminderSubject_(task);
    var body = buildReminderBody_(tester, task, activity, null);

    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body
    });

    recordReminderSentWithoutLock_(activity.activityId);

    return {
      sent: true,
      to: recipient,
      isTestMode: isTestMode,
      testerId: tester.testerId,
      taskId: task.taskId,
      activityId: activity.activityId
    };
  });
}

function sendMorningReminders() {
  var summary = {
    processed: 0,
    sent: 0,
    skipped: 0,
    failed: 0,
    errors: []
  };

  getReminderEligibleTesters().forEach(function (tester) {
    summary.processed += 1;
    try {
      var result = sendMorningReminderForTester_(tester);
      if (result.sent) {
        summary.sent += 1;
      } else {
        summary.skipped += 1;
      }
    } catch (error) {
      summary.failed += 1;
      summary.errors.push({
        testerId: tester.testerId,
        message: error.message
      });
    }
  });

  return summary;
}

function sendEveningReminders() {
  var summary = {
    processed: 0,
    sent: 0,
    skipped: 0,
    failed: 0,
    errors: []
  };

  getReminderEligibleTesters().forEach(function (tester) {
    summary.processed += 1;
    try {
      var result = sendEveningReminderForTester_(tester);
      if (result.sent) {
        summary.sent += 1;
      } else {
        summary.skipped += 1;
      }
    } catch (error) {
      summary.failed += 1;
      summary.errors.push({
        testerId: tester.testerId,
        message: error.message
      });
    }
  });

  return summary;
}

function runMorningReminderTest() {
  return sendMorningReminders();
}

function runEveningReminderTest() {
  return sendEveningReminders();
}

function installReminderTriggers() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var existing = ScriptApp.getProjectTriggers();
    var hasMorning = existing.some(function (t) {
      return t.getHandlerFunction() === REMINDER_TRIGGER_HANDLERS.MORNING;
    });
    var hasEvening = existing.some(function (t) {
      return t.getHandlerFunction() === REMINDER_TRIGGER_HANDLERS.EVENING;
    });

    var result = {
      createdMorning: false,
      createdEvening: false,
      existingMorning: hasMorning,
      existingEvening: hasEvening
    };

    if (!hasMorning) {
      createDailyAutomationTrigger_(REMINDER_TRIGGER_HANDLERS.MORNING, 9);
      result.createdMorning = true;
    }

    if (!hasEvening) {
      createDailyAutomationTrigger_(REMINDER_TRIGGER_HANDLERS.EVENING, 18);
      result.createdEvening = true;
    }

    return result;
  } finally {
    lock.releaseLock();
  }
}

function removeReminderTriggers() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var triggers = ScriptApp.getProjectTriggers();
    var removed = 0;
    triggers.forEach(function (trigger) {
      var handler = trigger.getHandlerFunction();
      if (
        handler === REMINDER_TRIGGER_HANDLERS.MORNING ||
        handler === REMINDER_TRIGGER_HANDLERS.EVENING
      ) {
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

function runEmailAutomationSmokeTest() {
  if (!isEmailTestModeEnabled_()) {
    throw new Error('Email test mode must be enabled. Run setEmailTestMode(true) and setTestEmailRecipient("you@example.com") first.');
  }
  var testRecipient = getTestEmailRecipient();
  if (!testRecipient) {
    throw new Error('Test email recipient is not configured.');
  }

  // Ensure default tasks are present without duplicating.
  var firstSeed = seedDefaultTasks();
  var secondSeed = seedDefaultTasks();
  assertEmailSmokeTest_(secondSeed.created === 0, 'Task seeding created duplicates.');

  var smokeId = Utilities.getUuid().replace(/-/g, '');
  var email = 'email-smoke-' + smokeId + '@example.invalid';
  var tester = addTester('[Smoke Test] Email Automation Tester', email, '');
  var activated = activateTester(tester.testerId, new Date());
  var currentDay = getTesterCurrentDay(activated.testerId);
  assertEmailSmokeTest_(currentDay === 1, 'Smoke tester should be on Day 1.');

  var today = getTodayTasksForTester(activated.testerId);
  assertEmailSmokeTest_(today.tasks.length > 0, 'No tasks found for today.');
  var amTask = getTaskForDayAndPeriod(currentDay, TASK_PERIODS.AM);
  var pmTask = getTaskForDayAndPeriod(currentDay, TASK_PERIODS.PM);
  assertEmailSmokeTest_(amTask !== null, 'AM task missing.');
  assertEmailSmokeTest_(pmTask !== null, 'PM task missing.');

  // Ensure activities are assigned.
  var amActivityBefore = assignTaskToTester(activated.testerId, amTask.taskId);
  var pmActivityBefore = assignTaskToTester(activated.testerId, pmTask.taskId);
  var amFresh = getActivityForTesterAndTask(activated.testerId, amTask.taskId);
  assertEmailSmokeTest_(amFresh.reminderCount === 0, 'New activity reminder count should be zero.');

  // Verify recipient redirection.
  var resolved = resolveEmailRecipient_(activated.email);
  assertEmailSmokeTest_(resolved === testRecipient, 'Email should be redirected to test recipient in test mode.');
  assertEmailSmokeTest_(resolved !== activated.email, 'Smoke test must not use real tester email in test mode.');

  // Morning email should send once.
  var morningResult = sendMorningReminderForTester_(activated);
  assertEmailSmokeTest_(morningResult.sent === true, 'Morning reminder should have been sent.');
  var amAfterMorning = getActivityForTesterAndTask(activated.testerId, amTask.taskId);
  assertEmailSmokeTest_(amAfterMorning.reminderCount === amFresh.reminderCount + 1, 'Reminder count should increment after morning send.');
  assertEmailSmokeTest_(amAfterMorning.lastReminderAt instanceof Date, 'Last Reminder At should be set.');

  // Duplicate morning protection.
  var morningDuplicate = sendMorningReminderForTester_(activated);
  assertEmailSmokeTest_(
    morningDuplicate.sent === false && morningDuplicate.reason === 'already_reminded',
    'Duplicate morning reminder should be blocked.'
  );
  var amAfterDuplicate = getActivityForTesterAndTask(activated.testerId, amTask.taskId);
  assertEmailSmokeTest_(amAfterDuplicate.reminderCount === amAfterMorning.reminderCount, 'Duplicate send should not increment count.');

  // Evening email should send once (consolidated).
  var eveningResult = sendEveningReminderForTester_(activated);
  assertEmailSmokeTest_(eveningResult.sent === true, 'Evening reminder should have been sent.');
  var pmAfterEvening = getActivityForTesterAndTask(activated.testerId, pmTask.taskId);
  assertEmailSmokeTest_(pmAfterEvening.reminderCount === pmActivityBefore.reminderCount + 1, 'PM reminder count should increment.');

  // Duplicate evening protection.
  var eveningDuplicate = sendEveningReminderForTester_(activated);
  assertEmailSmokeTest_(
    eveningDuplicate.sent === false && eveningDuplicate.reason === 'already_reminded',
    'Duplicate evening reminder should be blocked.'
  );

  // Completed tasks are skipped.
  completeActivity(amAfterMorning.activityId);
  var completedMorning = sendMorningReminderForTester_(activated);
  assertEmailSmokeTest_(
    completedMorning.sent === false && completedMorning.reason === 'already_completed',
    'Completed AM task should be skipped.'
  );

  return {
    ok: true,
    testerId: activated.testerId,
    testRecipient: testRecipient,
    seedCreated: firstSeed.created,
    morningSent: morningResult.sent,
    eveningSent: eveningResult.sent,
    amReminderCount: getActivityForTesterAndTask(activated.testerId, amTask.taskId).reminderCount,
    pmReminderCount: getActivityForTesterAndTask(activated.testerId, pmTask.taskId).reminderCount,
    message: 'Smoke-test tester and activity rows remain for manual review. No real tester was emailed because test mode was enabled.'
  };
}

// Internal helpers for per-tester reminder flows.

function sendMorningReminderForTester_(tester) {
  return runWithReminderDeliveryLocks_(function () {
    tester = getTesterById(tester.testerId);
    if (!tester || REMINDER_ELIGIBLE_STATUSES.indexOf(tester.status) === -1) {
      return { sent: false, reason: 'not_eligible' };
    }
    var currentDay = getTesterCurrentDay(tester);
    if (currentDay < 1 || currentDay > TESTING_PLAN_DAYS) {
      return { sent: false, reason: 'out_of_range', currentDay: currentDay };
    }
    var task = getTaskForDayAndPeriod(currentDay, TASK_PERIODS.AM);
    if (!task) {
      return { sent: false, reason: 'no_task', currentDay: currentDay };
    }
    var activity = assignTaskToTesterWithResultWithoutLock_(
      tester.testerId,
      task.taskId
    ).activity;
    if (activity.status === ACTIVITY_STATUSES.COMPLETED) {
      return { sent: false, reason: 'already_completed', activityId: activity.activityId };
    }
    if (isAlreadyRemindedToday_(activity)) {
      return { sent: false, reason: 'already_reminded', activityId: activity.activityId };
    }

    var recipient = resolveEmailRecipient_(tester.email);
    var subject = buildReminderSubject_(task);
    var body = buildMorningBody_(tester, task, currentDay, activity);

    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body
    });

    recordReminderSentWithoutLock_(activity.activityId);

    return {
      sent: true,
      reason: 'sent',
      activityId: activity.activityId,
      taskId: task.taskId,
      to: recipient
    };
  });
}

function sendEveningReminderForTester_(tester) {
  return runWithReminderDeliveryLocks_(function () {
    tester = getTesterById(tester.testerId);
    if (!tester || REMINDER_ELIGIBLE_STATUSES.indexOf(tester.status) === -1) {
      return { sent: false, reason: 'not_eligible' };
    }
    var currentDay = getTesterCurrentDay(tester);
    if (currentDay < 1 || currentDay > TESTING_PLAN_DAYS) {
      return { sent: false, reason: 'out_of_range', currentDay: currentDay };
    }
    var pmTask = getTaskForDayAndPeriod(currentDay, TASK_PERIODS.PM);
    if (!pmTask) {
      return { sent: false, reason: 'no_task', currentDay: currentDay };
    }
    var pmActivity = assignTaskToTesterWithResultWithoutLock_(
      tester.testerId,
      pmTask.taskId
    ).activity;
    if (pmActivity.status === ACTIVITY_STATUSES.COMPLETED) {
      return { sent: false, reason: 'already_completed', activityId: pmActivity.activityId };
    }
    if (isAlreadyRemindedToday_(pmActivity)) {
      return { sent: false, reason: 'already_reminded', activityId: pmActivity.activityId };
    }

    var amTask = getTaskForDayAndPeriod(currentDay, TASK_PERIODS.AM);
    var amActivity = null;
    if (amTask) {
      amActivity = assignTaskToTesterWithResultWithoutLock_(
        tester.testerId,
        amTask.taskId
      ).activity;
    }

    var recipient = resolveEmailRecipient_(tester.email);
    var subject = buildReminderSubject_(pmTask);
    var body = buildEveningBody_(tester, pmTask, amActivity, currentDay, pmActivity);

    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body
    });

    recordReminderSentWithoutLock_(pmActivity.activityId);

    return {
      sent: true,
      reason: 'sent',
      activityId: pmActivity.activityId,
      taskId: pmTask.taskId,
      to: recipient
    };
  });
}

function buildReminderSubject_(task) {
  var day = task.day;
  if (task.period === TASK_PERIODS.AM) {
    return 'Lodge Manager Test \u2014 Day ' + day + ' Morning Activity';
  }
  return 'Lodge Manager Test \u2014 Day ' + day + ' Evening Check';
}

function buildMorningBody_(tester, task, day, activity) {
  var name = tester.name || 'there';
  var feedbackUrl = getFeedbackUrlForActivity(activity.activityId);
  var lines = [
    'Hi ' + name + ',',
    '',
    'Your Day ' + day + ' Lodge Manager test is ready.',
    '',
    "Today's activity:",
    task.title,
    '',
    task.instructions,
    '',
    'Please open Lodge Manager and complete the activity.',
    '',
    'This should only take a few minutes.',
    '',
    'Complete the activity, then submit your feedback here:',
    feedbackUrl,
    '',
    'Thank you for helping us test Lodge Manager.'
  ];
  addTestModeNotice_(lines);
  return lines.join('\n');
}

function buildEveningBody_(tester, pmTask, amActivity, day, pmActivity) {
  var name = tester.name || 'there';
  var feedbackUrl = getFeedbackUrlForActivity(pmActivity.activityId);
  var lines = [
    'Hi ' + name + ',',
    '',
    'A quick final check for Day ' + day + '.',
    '',
    "Tonight's activity:",
    pmTask.title,
    '',
    pmTask.instructions,
    '',
    'Please open Lodge Manager and complete the activity if you have not already.',
    ''
  ];

  if (amActivity) {
    if (amActivity.status === ACTIVITY_STATUSES.COMPLETED) {
      lines.push('Morning activity: Completed \u2705');
      lines.push('');
    } else {
      lines.push('Your morning activity is still outstanding.');
      lines.push('');
    }
  }

  lines.push('Complete the activity, then submit your feedback here:');
  lines.push(feedbackUrl);
  lines.push('');
  lines.push("Thank you for today's testing.");
  addTestModeNotice_(lines);

  return lines.join('\n');
}

function buildReminderBody_(tester, task, activity, amActivity) {
  var day = task.day;
  if (task.period === TASK_PERIODS.AM) {
    return buildMorningBody_(tester, task, day, activity);
  }
  return buildEveningBody_(tester, task, amActivity, day, activity);
}

function addTestModeNotice_(lines) {
  if (isEmailTestModeEnabled_()) {
    lines.push('');
    lines.push('[TEST MODE] This message was redirected to the configured test recipient.');
  }
}

function runWithReminderDeliveryLocks_(action) {
  var scriptLock = LockService.getScriptLock();
  var documentLock = LockService.getDocumentLock();
  scriptLock.waitLock(30000);

  try {
    documentLock.waitLock(30000);
    try {
      return action();
    } finally {
      documentLock.releaseLock();
    }
  } finally {
    scriptLock.releaseLock();
  }
}

function assertEmailSmokeTest_(condition, message) {
  if (!condition) {
    throw new Error('Email automation smoke test failed: ' + message);
  }
}
