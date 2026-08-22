function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tester Automation')
    .addItem('Initialize Spreadsheet', 'initializeSpreadsheet')
    .addItem('Run Health Check', 'healthCheck')
    .addSeparator()
    .addItem('Run Tester Smoke Test', 'runTesterServiceSmokeTest')
    .addSeparator()
    .addItem('Seed Default Tasks', 'seedDefaultTasks')
    .addItem('Assign Today\'s Tasks for Active Testers', 'assignTodayTasksForAllActiveTesters')
    .addItem('Run Task Engine Smoke Test', 'runTaskEngineSmokeTest')
    .addToUi();
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
