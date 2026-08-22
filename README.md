# Tester Automation

Iterations 1 through 3 establish the local Google Apps Script foundation, tester management, and a manual 14-day task engine for a Play Store closed-test tester automation system. The project is container-bound to its target Google Spreadsheet, which is its datastore. No scheduling, reminders, emails, feedback forms, web handlers, or dashboard calculations are included.

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

## Tester Management

Iteration 2 manages testers only. All functions use the bound spreadsheet through `SpreadsheetApp.getActiveSpreadsheet()`.

### Testers Sheet

The `Testers` sheet must retain these headers in this order:

| Header | Purpose |
| --- | --- |
| Tester ID | UUID assigned when a tester is created. |
| Name | Tester name. |
| Email | Required, normalized to lowercase, and unique without regard to case. |
| Start Date | Optional until a tester is activated. |
| Status | One of the supported tester statuses. |
| Token | A private random identifier reserved for future personalized links. |
| Created At | Creation timestamp. |
| Updated At | Timestamp refreshed by every tester update. |

### Add And Find Testers

Run functions from the bound Apps Script editor after `initializeSpreadsheet()` has created the sheet:

```javascript
addTester('John Doe', 'john@example.com', '2026-08-21');
getTesterById('TESTER_ID');
getTesterByEmail('john@example.com');
getTesterByToken('TESTER_TOKEN');
getAllTesters();
getActiveTesters();
```

`addTester(name, email, startDate)` requires a non-empty name and a valid email. The start date may be blank, for example `addTester('John Doe', 'john@example.com', '')`.

Tester IDs use `Utilities.getUuid()` rather than sequential values. Tokens combine two random UUID values and are intended to be difficult to guess. Tokens are returned by tester lookup functions for application use, but the service does not log them.

### Statuses And Updates

Supported statuses are:

- `Not Started`
- `Active`
- `Needs Reminder`
- `At Risk`
- `Completed`
- `Inactive`

Update only editable fields with lower-camel-case keys:

```javascript
updateTester('TESTER_ID', {
  name: 'John Smith',
  email: 'john.smith@example.com',
  startDate: '2026-08-21',
  status: TESTER_STATUSES.ACTIVE
});
```

`Tester ID`, `Token`, and `Created At` cannot be changed. Email changes are validated and checked for duplicates. Every update refreshes `Updated At`; no status is calculated automatically.

To activate a tester manually, run:

```javascript
activateTester('TESTER_ID', '2026-08-21');
```

When the second argument is omitted or blank, `activateTester` uses the current date.

### Smoke Test

Run `runTesterServiceSmokeTest()` from the Apps Script editor after initialization. It creates one uniquely identified `[Smoke Test]` tester using an `example.invalid` email, verifies all lookup methods, validates an update and activation, and verifies duplicate-email prevention.

The smoke-test row is deliberately not deleted. It is safe to run repeatedly because each run uses a unique email. Review or manually remove only rows whose name starts with `[Smoke Test]` if they are no longer needed.

## 14-Day Task Engine

Iteration 3 defines a lightweight manual testing plan in the `Tasks` sheet. Each day has one `AM` task and one `PM` task, for 28 default tasks in total. Task records contain a UUID Task ID, Day, Period, Title, Instructions, and an Active flag.

### Default Task Plan

| Day | AM | PM |
| --- | --- | --- |
| 1 | Open Lodge Manager and sign in | Reopen Lodge Manager |
| 2 | Inspect the tenant list | Revisit a tenant record |
| 3 | Add a test tenant | Confirm the test tenant was saved |
| 4 | Review a tenancy record | Revisit tenancy details |
| 5 | Inspect reminders | Reopen reminders |
| 6 | Edit test tenant information | Confirm the edited information |
| 7 | Navigate between major screens | Try another navigation path |
| 8 | Review tenant and tenancy information | Revisit earlier test data |
| 9 | Return after a break | Repeat a familiar task |
| 10 | Inspect another tenant | Revisit a major screen |
| 11 | Create or update safe test data | Confirm saved test data |
| 12 | Review reminders and tenant details | Check a detail view again |
| 13 | Practice a typical app journey | Repeat the journey another way |
| 14 | Final beta app review | Final usability revisit |

Run `seedDefaultTasks()` once to add missing default Day/Period tasks. It is idempotent: an existing task for the same Day and Period is preserved, not overwritten, and not duplicated.

### Current Testing Day

`getTesterCurrentDay(testerOrTesterId)` requires an existing Active tester with a Start Date. The Start Date is Day 1. The calculation compares calendar dates in the Apps Script project timezone, not elapsed hours: it returns `0` before the Start Date, `1` through `14` during the plan, and `15` after the plan ends.

`getTodayTasksForTester(testerId)` returns the tester's current day and active AM/PM tasks. It returns no tasks before Day 1 or after Day 14. Inactive and not-started testers produce a clear error rather than receiving tasks.

### Assignments And Activities

The `Activity Log` records each assignment with an Activity ID, Tester ID, Task ID, Assigned At, Completed At, and Status. Supported activity statuses are:

- `Assigned`
- `Completed`
- `Skipped`

Use these manual functions from the Apps Script editor when needed:

```javascript
getAllTasks();
getTaskById('TASK_ID');
getTasksForDay(3);
getTaskForDayAndPeriod(3, 'AM');
assignTaskToTester('TESTER_ID', 'TASK_ID');
assignTodayTasksForTester('TESTER_ID');
assignTodayTasksForAllActiveTesters();
getActivityById('ACTIVITY_ID');
getActivitiesForTester('TESTER_ID');
getActivityForTesterAndTask('TESTER_ID', 'TASK_ID');
completeActivity('ACTIVITY_ID');
```

An assignment is unique for a Tester ID and Task ID. Reassigning the same pair returns the existing activity instead of adding a duplicate. `completeActivity(activityId)` is idempotent and preserves the first completion timestamp when it is called again.

### Spreadsheet Menu

After reloading the spreadsheet, the **Tester Automation** menu includes:

- Initialize Spreadsheet
- Run Health Check
- Run Tester Smoke Test
- Seed Default Tasks
- Assign Today's Tasks for Active Testers
- Run Task Engine Smoke Test

The assignment menu item processes all Active testers manually. It does not create a time-driven trigger.

### Task Engine Smoke Test

Run `Run Task Engine Smoke Test` from the spreadsheet menu, or execute `runTaskEngineSmokeTest()` in the Apps Script editor. It seeds missing default tasks, checks a repeat seed does not create duplicates, creates and activates a uniquely named `[Smoke Test] Task Engine Tester`, assigns Day 1 tasks, prevents a duplicate assignment, and marks one smoke-test activity completed.

The smoke-test tester and activity rows remain for review. It never deletes or changes existing real tester rows. Manually remove only rows clearly identified as `[Smoke Test]` if they are no longer needed.

## Local Development Workflow

1. Edit the relevant file in `src/` or `appsscript.json` locally.
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
