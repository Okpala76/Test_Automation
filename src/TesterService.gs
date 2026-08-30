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
      'Updated At': now,
      'Phone': '',
      'WhatsApp Enabled': false
    });

    context.sheet
      .getRange(context.sheet.getLastRow() + 1, 1, 1, context.headers.length)
      .setValues([row]);

    return testerObjectFromRow_(row, context.columnIndexes);
  } finally {
    lock.releaseLock();
  }
}

function bulkImportTesters() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var context = getTesterSheetContext_();
    var summary = {
      processed: 0,
      imported: 0,
      skipped: 0,
      failed: 0,
      errors: []
    };
    var lastRow = context.sheet.getLastRow();
    if (lastRow <= 1) {
      return summary;
    }

    var rows = context.sheet
      .getRange(2, 1, lastRow - 1, context.headers.length)
      .getValues();
    var testerIdIndex = context.columnIndexes['Tester ID'] - 1;
    var nameIndex = context.columnIndexes['Name'] - 1;
    var emailIndex = context.columnIndexes['Email'] - 1;
    var startDateIndex = context.columnIndexes['Start Date'] - 1;
    var statusIndex = context.columnIndexes['Status'] - 1;
    var tokenIndex = context.columnIndexes['Token'] - 1;
    var createdAtIndex = context.columnIndexes['Created At'] - 1;
    var updatedAtIndex = context.columnIndexes['Updated At'] - 1;
    var phoneIndex = context.columnIndexes['Phone'] - 1;
    var whatsappEnabledIndex = context.columnIndexes['WhatsApp Enabled'] - 1;
    var seenEmails = Object.create(null);
    var usedTesterIds = Object.create(null);
    var usedTokens = Object.create(null);

    rows.forEach(function (row) {
      var testerId = normalizeLookupValue_(row[testerIdIndex]);
      var token = normalizeLookupValue_(row[tokenIndex]);
      if (testerId) {
        usedTesterIds[testerId] = true;
        var existingEmail = normalizeLookupValue_(row[emailIndex]).toLowerCase();
        if (existingEmail) {
          seenEmails[existingEmail] = true;
        }
      } else if (
        normalizeLookupValue_(row[startDateIndex]) ||
        normalizeLookupValue_(row[statusIndex]) ||
        normalizeLookupValue_(row[tokenIndex]) ||
        normalizeLookupValue_(row[createdAtIndex]) ||
        normalizeLookupValue_(row[updatedAtIndex])
      ) {
        try {
          seenEmails[normalizeTesterEmail_(row[emailIndex])] = true;
        } catch (error) {
          // Invalid partial rows are reported during normal row processing.
        }
      }
      if (token) {
        usedTokens[token] = true;
      }
    });

    rows.forEach(function (row, index) {
      var sheetRow = index + 2;
      var testerId = normalizeLookupValue_(row[testerIdIndex]);
      var rawName = row[nameIndex];
      var rawEmail = row[emailIndex];
      if (
        testerId ||
        (!normalizeLookupValue_(rawName) && !normalizeLookupValue_(rawEmail))
      ) {
        return;
      }

      summary.processed += 1;
      try {
        normalizeTesterName_(rawName);
        var normalizedEmail = normalizeTesterEmail_(rawEmail);

        if (
          normalizeLookupValue_(row[startDateIndex]) ||
          normalizeLookupValue_(row[statusIndex]) ||
          normalizeLookupValue_(row[tokenIndex]) ||
          normalizeLookupValue_(row[createdAtIndex]) ||
          normalizeLookupValue_(row[updatedAtIndex])
        ) {
          summary.skipped += 1;
          summary.errors.push({
            row: sheetRow,
            type: 'skipped',
            message: 'Row contains existing onboarding data and was not overwritten.'
          });
          seenEmails[normalizedEmail] = true;
          return;
        }
        if (seenEmails[normalizedEmail]) {
          summary.skipped += 1;
          summary.errors.push({
            row: sheetRow,
            type: 'skipped',
            message: 'Duplicate tester email.'
          });
          return;
        }

        var newTesterId = generateUniqueBulkTesterValue_(
          generateTesterId,
          usedTesterIds
        );
        var newToken = generateUniqueBulkTesterValue_(
          generateTesterToken,
          usedTokens
        );
        var now = new Date();
        var importedRow = row.slice();
        importedRow[phoneIndex] = normalizeWhatsAppPhone(row[phoneIndex]);
        importedRow[whatsappEnabledIndex] = normalizeWhatsAppEnabled_(
          row[whatsappEnabledIndex]
        );
        importedRow[testerIdIndex] = newTesterId;
        importedRow[statusIndex] = TESTER_STATUSES.NOT_STARTED;
        importedRow[tokenIndex] = newToken;
        importedRow[createdAtIndex] = now;
        importedRow[updatedAtIndex] = now;

        context.sheet
          .getRange(sheetRow, 1, 1, context.headers.length)
          .setValues([importedRow]);
        usedTesterIds[newTesterId] = true;
        usedTokens[newToken] = true;
        seenEmails[normalizedEmail] = true;
        summary.imported += 1;
      } catch (error) {
        summary.failed += 1;
        summary.errors.push({
          row: sheetRow,
          type: 'failed',
          message: error.message
        });
      }
    });

    return summary;
  } finally {
    lock.releaseLock();
  }
}

function activateAllNotStartedTesters(startDate) {
  var normalizedStartDate = normalizeBulkActivationStartDate_(startDate);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var context = getTesterSheetContext_();
    var summary = {
      processed: 0,
      activated: 0,
      failed: 0,
      errors: [],
      startDate: normalizedStartDate
    };
    var lastRow = context.sheet.getLastRow();
    if (lastRow <= 1) {
      return summary;
    }

    var rows = context.sheet
      .getRange(2, 1, lastRow - 1, context.headers.length)
      .getValues();
    var testerIdIndex = context.columnIndexes['Tester ID'] - 1;
    var statusIndex = context.columnIndexes['Status'] - 1;
    rows.forEach(function (row, index) {
      if (
        !normalizeLookupValue_(row[testerIdIndex]) ||
        row[statusIndex] !== TESTER_STATUSES.NOT_STARTED
      ) {
        return;
      }

      var sheetRow = index + 2;
      summary.processed += 1;
      try {
        var now = new Date();
        var activatedRow = row.slice();
        activatedRow[context.columnIndexes['Start Date'] - 1] = normalizedStartDate;
        activatedRow[statusIndex] = TESTER_STATUSES.ACTIVE;
        activatedRow[context.columnIndexes['Updated At'] - 1] = now;
        context.sheet
          .getRange(sheetRow, 1, 1, context.headers.length)
          .setValues([activatedRow]);
        summary.activated += 1;
      } catch (error) {
        summary.failed += 1;
        summary.errors.push({
          row: sheetRow,
          message: error.message
        });
      }
    });

    return summary;
  } finally {
    lock.releaseLock();
  }
}

function countNotStartedTesters_() {
  return getAllTesters().filter(function (tester) {
    return tester.status === TESTER_STATUSES.NOT_STARTED;
  }).length;
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

  if (Object.prototype.hasOwnProperty.call(updates, 'phone')) {
    validatedUpdates.phone = normalizeWhatsAppPhone(updates.phone);
  }

  if (Object.prototype.hasOwnProperty.call(updates, 'whatsappEnabled')) {
    validatedUpdates.whatsappEnabled = normalizeWhatsAppEnabled_(updates.whatsappEnabled);
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

  if (Object.prototype.hasOwnProperty.call(validatedUpdates, 'phone')) {
    match.sheet
      .getRange(match.rowNumber, match.columnIndexes['Phone'])
      .setValue(validatedUpdates.phone);
  }

  if (Object.prototype.hasOwnProperty.call(validatedUpdates, 'whatsappEnabled')) {
    match.sheet
      .getRange(match.rowNumber, match.columnIndexes['WhatsApp Enabled'])
      .setValue(validatedUpdates.whatsappEnabled);
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
  ensureTesterWhatsAppColumns_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TESTER_SHEET_NAME);
  var requiredHeaders = REQUIRED_SHEET_HEADERS[TESTER_SHEET_NAME];

  if (!sheet) {
    throw new Error('The Testers sheet is missing. Run initializeSpreadsheet() first.');
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    throw new Error('The Testers sheet headers are invalid. Run initializeSpreadsheet() first.');
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (header) {
    return String(header).trim();
  });
  requiredHeaders.forEach(function (header) {
    if (headers.filter(function (candidate) { return candidate === header; }).length !== 1) {
      throw new Error(
        'The Testers sheet headers are invalid. Run initializeSpreadsheet() or restore the required headers.'
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

function countTesterRowsById_(testerId) {
  var context = getTesterSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return 0;
  }
  var idColumn = context.columnIndexes['Tester ID'];
  return context.sheet
    .getRange(2, idColumn, lastRow - 1, 1)
    .getValues()
    .filter(function (row) {
      return normalizeLookupValue_(row[0]) === testerId;
    }).length;
}

function buildTesterRow_(headers, valuesByHeader) {
  return headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(valuesByHeader, header)
      ? valuesByHeader[header]
      : '';
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
    updatedAt: row[columnIndexes['Updated At'] - 1],
    phone: normalizeLookupValue_(row[columnIndexes['Phone'] - 1]),
    whatsappEnabled: normalizeWhatsAppEnabledForRead_(
      row[columnIndexes['WhatsApp Enabled'] - 1]
    )
  };
}

function ensureTesterWhatsAppColumns_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TESTER_SHEET_NAME);
  if (!sheet || sheet.getLastColumn() === 0) {
    return;
  }
  var required = REQUIRED_SHEET_HEADERS[TESTER_SHEET_NAME];
  var lastCol = sheet.getLastColumn();
  var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (value) {
    return String(value).trim();
  });
  var missing = required.filter(function (header) {
    return currentHeaders.indexOf(header) === -1;
  });
  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
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

function generateUniqueBulkTesterValue_(generator, usedValues) {
  for (var attempt = 0; attempt < 10; attempt += 1) {
    var value = generator();
    if (!Object.prototype.hasOwnProperty.call(usedValues, value)) {
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

function normalizeBulkActivationStartDate_(startDate) {
  if (startDate instanceof Date) {
    var normalizedDate = normalizeTesterStartDate_(startDate);
    if (!normalizedDate) {
      throw new Error('A Start Date is required.');
    }
    return normalizedDate;
  }
  if (typeof startDate !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(startDate.trim())) {
    throw new Error('Start Date must use YYYY-MM-DD format.');
  }

  var input = startDate.trim();
  var timezone = Session.getScriptTimeZone();
  var parsed;
  try {
    parsed = Utilities.parseDate(input, timezone, 'yyyy-MM-dd');
  } catch (error) {
    throw new Error('Start Date must be a valid calendar date in YYYY-MM-DD format.');
  }
  if (Utilities.formatDate(parsed, timezone, 'yyyy-MM-dd') !== input) {
    throw new Error('Start Date must be a valid calendar date in YYYY-MM-DD format.');
  }
  return parsed;
}

function normalizeWhatsAppPhone(phone) {
  if (phone === null || typeof phone === 'undefined') {
    return '';
  }
  var input = String(phone).trim();
  if (!input) {
    return '';
  }
  var normalized = input.replace(/[+\s()\-]/g, '');
  if (!/^\d+$/.test(normalized)) {
    throw new Error('Phone must contain only an international country code and digits.');
  }
  if (!/^[1-9]\d{6,14}$/.test(normalized)) {
    throw new Error(
      'Phone must be an international number with 7 to 15 digits and cannot start with 0.'
    );
  }
  return normalized;
}

function normalizeWhatsAppEnabled_(value) {
  if (value === true || value === 1) {
    return true;
  }
  if (
    value === false ||
    value === 0 ||
    value === null ||
    typeof value === 'undefined' ||
    String(value).trim() === ''
  ) {
    return false;
  }
  var normalized = String(value).trim().toLowerCase();
  if (['true', 'yes', 'y', '1', 'on'].indexOf(normalized) !== -1) {
    return true;
  }
  if (['false', 'no', 'n', '0', 'off'].indexOf(normalized) !== -1) {
    return false;
  }
  throw new Error('WhatsApp Enabled must be TRUE or FALSE.');
}

function normalizeWhatsAppEnabledForRead_(value) {
  try {
    return normalizeWhatsAppEnabled_(value);
  } catch (error) {
    return false;
  }
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
