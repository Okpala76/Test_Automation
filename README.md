# Tester Automation

Iteration 1 establishes the local Google Apps Script foundation for a Play Store closed-test tester automation system. The project is container-bound to its target Google Spreadsheet, which is its eventual datastore. No scheduling, reminders, forms, dashboards, web handlers, or 14-day logic are included.

## Prerequisites

- Node.js and npm
- Git
- A Google account that can edit the target Google Spreadsheet and its container-bound Apps Script project
- The Google Apps Script API enabled for that Google account at https://script.google.com/home/usersettings

## Install

Install the project dependencies:

```sh
npm install
```

`@google/clasp` is installed locally as a development dependency, so all npm scripts use the pinned project version.

## Authenticate clasp

Run:

```sh
npm run clasp:login
```

Complete the Google sign-in and authorization flow in the browser using the account that will own or access the Apps Script project. clasp stores OAuth credentials outside this repository; `.clasprc.json` files are ignored by Git.

## Link The Bound Script

These steps require manual Google account interaction and cannot be completed locally without it.

1. Open the target Google Spreadsheet, then open its container-bound Apps Script project through **Extensions > Apps Script**.
2. Confirm that the Script ID of that project is already present in `.clasp.json`. This Script ID is only for clasp synchronization.
3. Authenticate clasp with `npm run clasp:login` if you have not already done so.
4. Push the local source with `npm run clasp:push`.

The project deliberately keeps `appsscript.json` at the repository root and Apps Script source in `src/`. Therefore `.clasp.json` uses `"rootDir": "."`, which tells clasp that the repository root is the content directory while preserving the `src/` paths.

## Initialize And Check The Spreadsheet

After pushing, run `initializeSpreadsheet()` from the Apps Script editor. It uses `SpreadsheetApp.getActiveSpreadsheet()` to access the spreadsheet containing this container-bound script. It creates any missing sheets and adds headers only to completely empty sheets. It can be run repeatedly without creating duplicate sheets or headers, deleting data, or overwriting entered data.

Google will ask you to authorize the script the first time it accesses Google Sheets. Review and complete that authorization in the Apps Script editor.

The expected sheets are:

- `Testers`
- `Tasks`
- `Activity Log`
- `Feedback`
- `Dashboard`

Run `healthCheck()` afterward. It returns and logs a structured result with `ok`, `spreadsheetName`, `requiredSheets`, and `missingSheets`. Before initialization, it reports exactly which of the five sheets are missing.

## Local Development Workflow

1. Edit `src/Code.gs`, `src/Config.gs`, or `appsscript.json` locally.
2. Run `npm run clasp:push` to upload the local project after it has been linked and authenticated.
3. Run Apps Script functions in the remote editor to authorize and verify behavior against the configured spreadsheet.
4. Run `npm run clasp:pull` only when remote Apps Script changes need to be brought back locally, then review the diff before continuing.
5. Use `npm run clasp:open` to open the linked Apps Script project.

Available npm scripts:

```sh
npm run clasp:login
npm run clasp:push
npm run clasp:pull
npm run clasp:open
```
