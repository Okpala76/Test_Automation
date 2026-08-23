function generateTesterId() {
  return Utilities.getUuid();
}

function generateTesterToken() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function addTester(name, email, startDate) {
  var normalizedName = normalizeTesterName_(name);
  var normalizedEmail = normalizeTesterEmail_(email);
  var normalizedStartDate = normalizeTesterStartDate_(startDate);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var context = getTesterSheetContext_();
    if (getTesterByEmail(normalizedEmail)) {
      throw new Error('A tester with this email already exists.');
    }

    var testerId = generateUniqueTesterValue_(generateTesterId, getTesterById);
    var token = generateUniqueTesterValue_(generateTesterToken, getTesterByToken);
    var now = new Date();
    var row = buildTesterRow_(context.headers, {
      'Tester ID': testerId,
      'Name': normalizedName,
      'Email': normalizedEmail,
      'Start Date': normalizedStartDate,
      'Status': TESTER_STATUSES.NOT_STARTED,
      'Token': token,
      'Created At': now,
      'Updated At': now
    });

    context.sheet
      .getRange(context.sheet.getLastRow() + 1, 1, 1, context.headers.length)
      .setValues([row]);

    return testerObjectFromRow_(row, context.columnIndexes);
  } finally {
    lock.releaseLock();
  }
}

function getTesterById(testerId) {
  var normalizedId = normalizeLookupValue_(testerId);
  if (!normalizedId) {
    return null;
  }

  var match = findTesterRow_('Tester ID', normalizedId);
  return match ? testerObjectFromRow_(match.row, match.columnIndexes) : null;
}

function getTesterByEmail(email) {
  var normalizedEmail = normalizeLookupValue_(email).toLowerCase();
  if (!normalizedEmail) {
    return null;
  }

  var match = findTesterRow_('Email', normalizedEmail, true);
  return match ? testerObjectFromRow_(match.row, match.columnIndexes) : null;
}

function getTesterByToken(token) {
  var normalizedToken = normalizeLookupValue_(token);
  if (!normalizedToken) {
    return null;
  }

  var match = findTesterRow_('Token', normalizedToken);
  return match ? testerObjectFromRow_(match.row, match.columnIndexes) : null;
}

function updateTester(testerId, updates) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    return updateTesterWithoutLock_(testerId, updates);
  } finally {
    lock.releaseLock();
  }
}

function updateTesterWithoutLock_(testerId, updates) {
  var normalizedId = normalizeLookupValue_(testerId);
  if (!normalizedId) {
    throw new Error('A tester ID is required.');
  }
  if (!updates || typeof updates !== 'object' || Array.isArray(updates)) {
    throw new Error('Updates must be an object containing editable tester fields.');
  }

  var updateKeys = Object.keys(updates);
  if (updateKeys.length === 0) {
    throw new Error('Provide at least one editable tester field to update.');
  }

  updateKeys.forEach(function (field) {
    if (!Object.prototype.hasOwnProperty.call(TESTER_EDITABLE_FIELDS, field)) {
      throw new Error('Field "' + field + '" cannot be updated.');
    }
  });

  var match = findTesterRow_('Tester ID', normalizedId);
  if (!match) {
    throw new Error('Tester not found for the provided tester ID.');
  }

  var validatedUpdates = {};
  if (Object.prototype.hasOwnProperty.call(updates, 'name')) {
    validatedUpdates.name = normalizeTesterName_(updates.name);
  }

  if (Object.prototype.hasOwnProperty.call(updates, 'email')) {
    var normalizedEmail = normalizeTesterEmail_(updates.email);
    var existingTester = getTesterByEmail(normalizedEmail);
    if (existingTester && existingTester.testerId !== normalizedId) {
      throw new Error('A tester with this email already exists.');
    }
    validatedUpdates.email = normalizedEmail;
  }

  if (Object.prototype.hasOwnProperty.call(updates, 'startDate')) {
    validatedUpdates.startDate = normalizeTesterStartDate_(updates.startDate);
  }

  if (Object.prototype.hasOwnProperty.call(updates, 'status')) {
    validatedUpdates.status = validateTesterStatus_(updates.status);
  }

  if (Object.prototype.hasOwnProperty.call(validatedUpdates, 'name')) {
    match.sheet
      .getRange(match.rowNumber, match.columnIndexes['Name'])
      .setValue(validatedUpdates.name);
  }

  if (Object.prototype.hasOwnProperty.call(validatedUpdates, 'email')) {
    match.sheet
      .getRange(match.rowNumber, match.columnIndexes['Email'])
      .setValue(validatedUpdates.email);
  }

  if (Object.prototype.hasOwnProperty.call(validatedUpdates, 'startDate')) {
    match.sheet
      .getRange(match.rowNumber, match.columnIndexes['Start Date'])
      .setValue(validatedUpdates.startDate);
  }

  if (Object.prototype.hasOwnProperty.call(validatedUpdates, 'status')) {
    match.sheet
      .getRange(match.rowNumber, match.columnIndexes['Status'])
      .setValue(validatedUpdates.status);
  }

  match.sheet
    .getRange(match.rowNumber, match.columnIndexes['Updated At'])
    .setValue(new Date());

  return getTesterById(normalizedId);
}

function getAllTesters() {
  var context = getTesterSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues()
    .filter(function (row) {
      return normalizeLookupValue_(row[context.columnIndexes['Tester ID'] - 1]) !== '';
    })
    .map(function (row) {
      return testerObjectFromRow_(row, context.columnIndexes);
    });
}

function getActiveTesters() {
  return getAllTesters().filter(function (tester) {
    return tester.status === TESTER_STATUSES.ACTIVE;
  });
}

function getReminderEligibleTesters() {
  return getAllTesters().filter(function (tester) {
    return REMINDER_ELIGIBLE_STATUSES.indexOf(tester.status) !== -1;
  });
}

function activateTester(testerId, startDate) {
  var activationDate =
    startDate === null || typeof startDate === 'undefined' || startDate === ''
      ? new Date()
      : startDate;

  return updateTester(testerId, {
    status: TESTER_STATUSES.ACTIVE,
    startDate: activationDate
  });
}

function runTesterServiceSmokeTest() {
  var smokeId = Utilities.getUuid().replace(/-/g, '');
  var email = 'tester-smoke-' + smokeId + '@example.invalid';
  var created = addTester('[Smoke Test] Temporary Tester', email, '');
  var byId = getTesterById(created.testerId);
  var byEmail = getTesterByEmail(email.toUpperCase());
  var byToken = getTesterByToken(created.token);

  assertTesterSmokeTest_(byId && byId.email === email, 'Lookup by tester ID failed.');
  assertTesterSmokeTest_(byEmail && byEmail.testerId === created.testerId, 'Lookup by email failed.');
  assertTesterSmokeTest_(byToken && byToken.testerId === created.testerId, 'Lookup by token failed.');

  var updated = updateTester(created.testerId, {
    name: '[Smoke Test] Updated Tester',
    status: TESTER_STATUSES.NEEDS_REMINDER
  });
  assertTesterSmokeTest_(
    updated.status === TESTER_STATUSES.NEEDS_REMINDER,
    'Tester update failed.'
  );

  var activated = activateTester(created.testerId);
  assertTesterSmokeTest_(
    activated.status === TESTER_STATUSES.ACTIVE && activated.startDate instanceof Date,
    'Tester activation failed.'
  );

  var duplicatePrevented = false;
  try {
    addTester('[Smoke Test] Duplicate Tester', email, '');
  } catch (error) {
    duplicatePrevented = true;
  }
  assertTesterSmokeTest_(duplicatePrevented, 'Duplicate email was not prevented.');

  return {
    ok: true,
    testerId: created.testerId,
    email: email,
    status: activated.status,
    message: 'Smoke-test data remains in the Testers sheet for manual review.'
  };
}

function getTesterSheetContext_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TESTER_SHEET_NAME);
  var headers = REQUIRED_SHEET_HEADERS[TESTER_SHEET_NAME];

  if (!sheet) {
    throw new Error('The Testers sheet is missing. Run initializeSpreadsheet() first.');
  }

  var sheetHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (header, index) {
    if (sheetHeaders[index] !== header) {
      throw new Error(
        'The Testers sheet headers are invalid. Run initializeSpreadsheet() on an empty sheet or restore the required headers.'
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

function findTesterRow_(columnName, value, ignoreCase) {
  var context = getTesterSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return null;
  }

  var columnIndex = context.columnIndexes[columnName];
  var values = context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues();
  var expectedValue = ignoreCase ? String(value).toLowerCase() : String(value);

  for (var index = 0; index < values.length; index += 1) {
    var candidate = normalizeLookupValue_(values[index][columnIndex - 1]);
    if (ignoreCase) {
      candidate = candidate.toLowerCase();
    }
    if (candidate === expectedValue) {
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

function buildTesterRow_(headers, valuesByHeader) {
  return headers.map(function (header) {
    return valuesByHeader[header];
  });
}

function testerObjectFromRow_(row, columnIndexes) {
  return {
    testerId: row[columnIndexes['Tester ID'] - 1],
    name: row[columnIndexes['Name'] - 1],
    email: row[columnIndexes['Email'] - 1],
    startDate: row[columnIndexes['Start Date'] - 1],
    status: row[columnIndexes['Status'] - 1],
    token: row[columnIndexes['Token'] - 1],
    createdAt: row[columnIndexes['Created At'] - 1],
    updatedAt: row[columnIndexes['Updated At'] - 1]
  };
}

function generateUniqueTesterValue_(generator, lookup) {
  for (var attempt = 0; attempt < 10; attempt += 1) {
    var value = generator();
    if (!lookup(value)) {
      return value;
    }
  }

  throw new Error('Unable to generate a unique tester identifier. Please try again.');
}

function normalizeTesterName_(name) {
  if (typeof name !== 'string' || name.trim() === '') {
    throw new Error('Tester name is required.');
  }

  return name.trim();
}

function normalizeTesterEmail_(email) {
  if (typeof email !== 'string') {
    throw new Error('A valid tester email is required.');
  }

  var normalizedEmail = email.trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizedEmail)) {
    throw new Error('A valid tester email is required.');
  }

  return normalizedEmail;
}

function normalizeTesterStartDate_(startDate) {
  if (
    startDate === null ||
    typeof startDate === 'undefined' ||
    (typeof startDate === 'string' && startDate.trim() === '')
  ) {
    return '';
  }

  var normalizedDate = startDate instanceof Date ? new Date(startDate.getTime()) : new Date(startDate);
  if (isNaN(normalizedDate.getTime())) {
    throw new Error('Start Date must be a valid date or blank.');
  }

  return normalizedDate;
}

function validateTesterStatus_(status) {
  if (Object.keys(TESTER_STATUSES).some(function (key) {
    return TESTER_STATUSES[key] === status;
  })) {
    return status;
  }

  throw new Error('Status must be one of the supported tester statuses.');
}

function normalizeLookupValue_(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }

  return String(value).trim();
}

function assertTesterSmokeTest_(condition, message) {
  if (!condition) {
    throw new Error('Tester service smoke test failed: ' + message);
  }
}
