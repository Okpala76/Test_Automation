function generateTaskId() {
  return Utilities.getUuid();
}

function seedDefaultTasks() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var context = getTaskSheetContext_();
    var existingPairs = {};
    getAllTasks().forEach(function (task) {
      existingPairs[getTaskDayPeriodKey_(task.day, task.period)] = true;
    });

    var rows = [];
    var skipped = 0;
    DEFAULT_TASKS.forEach(function (task) {
      var day = normalizeTaskDay_(task.day);
      var period = normalizeTaskPeriod_(task.period);
      var key = getTaskDayPeriodKey_(day, period);

      if (existingPairs[key]) {
        skipped += 1;
        return;
      }

      rows.push(
        buildTaskRow_(context.headers, {
          'Task ID': generateUniqueTaskId_(),
          'Day': day,
          'Period': period,
          'Title': task.title,
          'Instructions': task.instructions,
          'Active': true
        })
      );
      existingPairs[key] = true;
    });

    if (rows.length > 0) {
      context.sheet
        .getRange(context.sheet.getLastRow() + 1, 1, rows.length, context.headers.length)
        .setValues(rows);
    }

    return {
      created: rows.length,
      skipped: skipped
    };
  } finally {
    lock.releaseLock();
  }
}

function getAllTasks() {
  var context = getTaskSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues()
    .filter(function (row) {
      return normalizeTaskLookupValue_(row[context.columnIndexes['Task ID'] - 1]) !== '';
    })
    .map(function (row) {
      return taskObjectFromRow_(row, context.columnIndexes);
    });
}

function getTaskById(taskId) {
  var normalizedId = normalizeTaskLookupValue_(taskId);
  if (!normalizedId) {
    return null;
  }

  var match = findTaskRowById_(normalizedId);
  return match ? taskObjectFromRow_(match.row, match.columnIndexes) : null;
}

function getTasksForDay(day) {
  var normalizedDay = normalizeTaskDay_(day);
  return getAllTasks()
    .filter(function (task) {
      return task.day === normalizedDay && task.active;
    })
    .sort(sortTasksByPeriod_);
}

function getTaskForDayAndPeriod(day, period) {
  var normalizedDay = normalizeTaskDay_(day);
  var normalizedPeriod = normalizeTaskPeriod_(period);
  var tasks = getAllTasks();

  for (var index = 0; index < tasks.length; index += 1) {
    if (tasks[index].day === normalizedDay && tasks[index].period === normalizedPeriod) {
      return tasks[index];
    }
  }

  return null;
}

function getTesterCurrentDay(testerOrTesterId) {
  return calculateTesterCurrentDay_(resolveActiveTester_(testerOrTesterId));
}

function getTodayTasksForTester(testerId) {
  var tester = resolveActiveTester_(testerId);
  var currentDay = calculateTesterCurrentDay_(tester);

  return {
    testerId: tester.testerId,
    currentDay: currentDay,
    tasks:
      currentDay < 1 || currentDay > TESTING_PLAN_DAYS ? [] : getTasksForDay(currentDay)
  };
}

function getTaskSheetContext_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SHEET_NAME);
  var headers = REQUIRED_SHEET_HEADERS[TASK_SHEET_NAME];

  if (!sheet) {
    throw new Error('The Tasks sheet is missing. Run initializeSpreadsheet() first.');
  }

  var sheetHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (header, index) {
    if (sheetHeaders[index] !== header) {
      throw new Error(
        'The Tasks sheet headers are invalid. Run initializeSpreadsheet() on an empty sheet or restore the required headers.'
      );
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

function findTaskRowById_(taskId) {
  var context = getTaskSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return null;
  }

  var values = context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues();
  var taskIdColumn = context.columnIndexes['Task ID'] - 1;

  for (var index = 0; index < values.length; index += 1) {
    if (normalizeTaskLookupValue_(values[index][taskIdColumn]) === taskId) {
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

function buildTaskRow_(headers, valuesByHeader) {
  return headers.map(function (header) {
    return valuesByHeader[header];
  });
}

function taskObjectFromRow_(row, columnIndexes) {
  return {
    taskId: normalizeTaskLookupValue_(row[columnIndexes['Task ID'] - 1]),
    day: normalizeTaskDay_(row[columnIndexes['Day'] - 1]),
    period: normalizeTaskPeriod_(row[columnIndexes['Period'] - 1]),
    title: row[columnIndexes['Title'] - 1],
    instructions: row[columnIndexes['Instructions'] - 1],
    active: isTaskActive_(row[columnIndexes['Active'] - 1])
  };
}

function generateUniqueTaskId_() {
  for (var attempt = 0; attempt < 10; attempt += 1) {
    var taskId = generateTaskId();
    if (!getTaskById(taskId)) {
      return taskId;
    }
  }

  throw new Error('Unable to generate a unique task ID. Please try again.');
}

function normalizeTaskDay_(day) {
  var normalizedDay = Number(day);
  if (
    !isFinite(normalizedDay) ||
    Math.floor(normalizedDay) !== normalizedDay ||
    normalizedDay < 1 ||
    normalizedDay > TESTING_PLAN_DAYS
  ) {
    throw new Error('Task Day must be an integer from 1 through ' + TESTING_PLAN_DAYS + '.');
  }

  return normalizedDay;
}

function normalizeTaskPeriod_(period) {
  var normalizedPeriod = normalizeTaskLookupValue_(period).toUpperCase();
  if (
    normalizedPeriod !== TASK_PERIODS.AM &&
    normalizedPeriod !== TASK_PERIODS.PM
  ) {
    throw new Error('Task Period must be AM or PM.');
  }

  return normalizedPeriod;
}

function isTaskActive_(value) {
  return value === true || String(value).toLowerCase() === 'true';
}

function getTaskDayPeriodKey_(day, period) {
  return String(day) + '|' + period;
}

function sortTasksByPeriod_(firstTask, secondTask) {
  return firstTask.period === secondTask.period
    ? 0
    : firstTask.period === TASK_PERIODS.AM
      ? -1
      : 1;
}

function resolveActiveTester_(testerOrTesterId) {
  var testerId =
    typeof testerOrTesterId === 'object' && testerOrTesterId
      ? testerOrTesterId.testerId
      : testerOrTesterId;
  var tester = getTesterById(testerId);

  if (!tester) {
    throw new Error('Tester not found for the provided tester ID.');
  }
  if (tester.status !== TESTER_STATUSES.ACTIVE) {
    throw new Error('Tester must have Active status to receive testing tasks.');
  }
  if (!normalizeTesterStartDate_(tester.startDate)) {
    throw new Error('Active tester must have a Start Date.');
  }

  return tester;
}

function calculateTesterCurrentDay_(tester) {
  var startDate = normalizeTesterStartDate_(tester.startDate);
  var startCalendarDate = calendarDateAsUtcMillis_(startDate);
  var todayCalendarDate = calendarDateAsUtcMillis_(new Date());
  var elapsedCalendarDays = Math.round(
    (todayCalendarDate - startCalendarDate) / (24 * 60 * 60 * 1000)
  );
  var currentDay = elapsedCalendarDays + 1;

  if (currentDay < 1) {
    return 0;
  }
  if (currentDay > TESTING_PLAN_DAYS) {
    return POST_TEST_DAY;
  }

  return currentDay;
}

function calendarDateAsUtcMillis_(date) {
  var timezone = Session.getScriptTimeZone();
  var parts = Utilities.formatDate(date, timezone, 'yyyy-MM-dd').split('-');

  // The date parts are normalized to midnight UTC solely for calendar-day comparison.
  return Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function normalizeTaskLookupValue_(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }

  return String(value).trim();
}
