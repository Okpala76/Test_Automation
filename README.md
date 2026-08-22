# Tester Automation

Iterations 1 through 6 establish the local Google Apps Script foundation, tester management, a 14-day task engine, automated email reminders, a tester-facing feedback Web App, and smart participation monitoring. The project remains container-bound to its target Google Spreadsheet. Final dashboard charts and escalation messaging are not included yet.

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
- `Monitoring`

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

The `Activity Log` records each assignment with an Activity ID, Tester ID, Task ID, Assigned At, Completed At, Status, Last Reminder At, and Reminder Count. Iteration 4 added `Last Reminder At` and `Reminder Count` to track reminder delivery without destroying existing data — existing sheets are migrated idempotently by appending the two new columns. Supported activity statuses are:

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
- Email Automation → Run Morning Reminders, Run Evening Reminders, Install Reminder Triggers, Remove Reminder Triggers, Run Email Automation Smoke Test
- Feedback → Configure Web App URL, Check Web App Configuration, Run Feedback Smoke Test, Show Deployment Instructions
- Monitoring → Refresh Monitoring, Refresh Tester Statuses, Install Monitoring Trigger, Remove Monitoring Trigger, Run Monitoring Smoke Test

Every menu action uses a spreadsheet toast for normal results and an alert for important errors. Return values and execution logs remain available for debugging, but are not required for normal operation.

### Task Engine Smoke Test

Run `Run Task Engine Smoke Test` from the spreadsheet menu, or execute `runTaskEngineSmokeTest()` in the Apps Script editor. It seeds missing default tasks, checks a repeat seed does not create duplicates, creates and activates a uniquely named `[Smoke Test] Task Engine Tester`, assigns Day 1 tasks, prevents a duplicate assignment, and marks one smoke-test activity completed.

The smoke-test tester and activity rows remain for review. It never deletes or changes existing real tester rows. Manually remove only rows clearly identified as `[Smoke Test]` if they are no longer needed.

## Email Reminder Automation

Iteration 4 automatically reminds eligible testers about their current Day 1–14 AM and PM tasks using `MailApp` and the bound spreadsheet as the source of truth. It reuses the existing current-day calculation and never sends reminders before a tester’s Start Date or after Day 14.

### Activity Log Reminder Tracking

The two new columns prevent repeated emails when a function is run twice or a trigger retries. `Last Reminder At` records when the last reminder was sent for that Tester/Task pair, and `Reminder Count` increments each time. `initializeSpreadsheet()` and `ensureActivityLogReminderColumns_()` append these headers idempotently without reordering existing data.

### Test Mode

Email delivery must never accidentally reach real testers during testing. Test mode redirects all reminder emails to a single developer address.

Configure it once in the bound Apps Script editor:

```javascript
setEmailTestMode(true);
setTestEmailRecipient('you@example.com');
```

Check the current configuration with:

```javascript
getEmailTestMode();
getTestEmailRecipient();
```

When test mode is enabled, every call to `sendTaskReminderEmail`, `sendMorningReminders`, and `sendEveningReminders` sends to the test recipient instead of the tester’s real email. Disable test mode for real runs with `setEmailTestMode(false)`. No personal email is hardcoded.

### Morning Reminders

`sendMorningReminders()` iterates over all reminder-eligible testers, determines each tester’s current day, assigns today’s AM task if needed, and skips testers that are out of range, already completed, or already reminded today. A successful send records `Last Reminder At` and increments `Reminder Count`.

Use the safe no-argument wrapper from the editor or menu:

```javascript
runMorningReminderTest(); // respects test mode
```

The underlying sender is:

```javascript
sendTaskReminderEmail(tester, task, activity);
```

The morning email includes tester name, day, AM designation, task title, instructions, and a personalized feedback link for the assigned activity. Evening emails include the equivalent PM feedback link. Test Mode still redirects the email to the configured test inbox while preserving the smoke tester's personalized link.

Example subject: `Lodge Manager Test — Day 4 Morning Activity`

### Evening Reminders

`sendEveningReminders()` sends one consolidated email per Active tester for today’s PM task. It assigns the PM task if missing, skips testers whose PM task is already completed or already reminded today, and includes a note about the morning activity:

- `Morning activity: Completed ✅` if the AM activity is completed
- `Your morning activity is still outstanding.` otherwise

Only one evening email is sent per tester per run.

Use:

```javascript
runEveningReminderTest();
```

### Duplicate Protection and Eligibility

- `Active`, `Needs Reminder`, and `At Risk` testers are eligible for their normal AM/PM task reminders.
- `Not Started`, `Completed`, and `Inactive` testers are not eligible for normal reminders.
- Tasks outside Day 1–14 are ignored.
- Only active tasks are assigned and reminded.
- Already completed activities are skipped (`reason: "already_completed"`).
- Already reminded activities for the same calendar day (script timezone) are skipped (`reason: "already_reminded"`), so running a reminder twice or a trigger retry does not send duplicate period emails.
- One failing tester/email does not stop batch processing; `sendMorningReminders()` and `sendEveningReminders()` return a summary like `{ processed, sent, skipped, failed, errors }` without exposing tokens.

### Triggers

Create daily time-driven triggers in the Apps Script project timezone:

```javascript
installReminderTriggers(); // creates ~9:00 AM and ~6:00 PM triggers
removeReminderTriggers(); // removes only the two reminder triggers
```

Both functions are idempotent: running `installReminderTriggers()` repeatedly does not create duplicates, and `removeReminderTriggers()` touches only handlers named `sendMorningReminders` and `sendEveningReminders`. Inspect triggers manually at **Triggers** in the Apps Script editor. Do not assume exact-minute execution.

Project timezone matters for both day calculation and trigger scheduling. Confirm it in **Project Settings → Time zone** (e.g., `Etc/UTC`).

### Email Automation Smoke Test

Run `Run Email Automation Smoke Test` from the menu, or execute `runEmailAutomationSmokeTest()` in the Apps Script editor with test mode enabled:

```javascript
setEmailTestMode(true);
setTestEmailRecipient('you@example.com');
runEmailAutomationSmokeTest();
```

The smoke test verifies: test mode is configured, default tasks can be seeded idempotently, a uniquely named `[Smoke Test] Email Automation Tester` is created and activated for today, morning and evening emails are generated and sent to the test recipient, duplicate reminder protection blocks a second send for the same period/day, completed tasks are skipped, reminder counts and timestamps update correctly, and no real tester receives a smoke-test email (all sends are redirected while test mode is on).

Smoke-test rows persist for manual review; only remove rows clearly marked `[Smoke Test]` if they are no longer needed.

## Tester Feedback Web App

Iteration 5 exposes a mobile-first feedback page without granting testers access to the spreadsheet. The URL contains only the tester's existing random bearer token and an Activity ID:

```text
https://script.google.com/macros/s/DEPLOYMENT_ID/exec?tester=TESTER_TOKEN&activity=ACTIVITY_ID
```

The URL does not expose email addresses, row numbers, Spreadsheet IDs, or Script IDs. `doGet(e)` validates the token, activity, activity ownership, and task before rendering any tester or task details. Invalid requests receive a friendly generic page without stack traces or internal data.

### Feedback Page

The page displays the tester's name, task day and AM/PM period, task title, and instructions. It collects:

- Whether the activity was completed
- Rating from 1 through 5
- Whether a bug was encountered
- An optional comment of up to 2,000 characters

The browser calls `submitFeedback(payload)` through `google.script.run`. All values are validated again on the server; URL and browser values are never trusted by themselves.

### Submission Behavior

Valid submissions append one row to the existing `Feedback` sheet using these fields:

```text
Feedback ID | Tester ID | Task ID | Rating | Completed | Bug Reported | Comment | Submitted At
```

`Completed` and `Bug Reported` are stored as Booleans. Tokens are never stored in `Feedback`. When `Completed` is true, submission reuses the existing activity completion logic so the Activity Log status becomes `Completed` and `Completed At` is set. A false completion answer leaves the activity incomplete while still recording feedback.

Feedback is idempotent under a document lock. Only one Feedback row is permitted for each Tester ID and Task ID. Repeated button clicks or browser retries return a successful `alreadySubmitted` result and do not append another row.

### Feedback URL Configuration

The Web App deployment URL is stored in Script Properties, never hardcoded:

```javascript
setFeedbackWebAppUrl('https://script.google.com/macros/s/DEPLOYMENT_ID/exec');
getFeedbackWebAppUrl();
getFeedbackUrlForActivity('ACTIVITY_ID');
```

Reminder emails require this configuration so they can include a personalized link. If it is missing, that tester's reminder fails safely and the batch summary reports the failure.

### Security Model

The tester token is treated as a bearer identifier. Keep personalized links private. The Web App returns only the single validated tester/task context needed by the page, never the full tester dataset. Every submission rechecks ownership server-side and permits no arbitrary row access. Testers do not need spreadsheet edit or view access.

### Feedback Smoke Test

After deploying and configuring the `/exec` URL, run **Tester Automation → Feedback → Run Feedback Smoke Test** or execute:

```javascript
runFeedbackSmokeTest();
```

The test creates a uniquely named `[Smoke Test] Feedback Web App Tester`, activates it on Day 1, assigns the Day 1 AM task, generates a personalized URL, validates token ownership, submits a five-star smoke-test Feedback row, completes the activity, and confirms a repeated submission is blocked. It also assigns the PM task and displays a fresh personalized URL in a spreadsheet alert for manual browser testing. It does not modify or delete real tester rows. Smoke-test rows remain for manual review.

### Web App Deployment

1. Push the current source with `npm run clasp:push`.
2. Open the bound Apps Script project with `npm run clasp:open` or **Extensions → Apps Script** from the spreadsheet.
3. Select **Deploy → New deployment**.
4. Click **Select type**, then choose **Web app**.
5. Set **Execute as** to the script owner so external testers do not need spreadsheet permission.
6. Set **Who has access** to the option that permits the intended external testers, typically **Anyone**. Google Workspace policy may restrict this option; do not proceed until the required tester access is available.
7. Deploy, complete the authorization prompt, and copy the Web App URL ending in `/exec`.
8. In the spreadsheet, choose **Tester Automation → Feedback → Configure Web App URL** and paste the `/exec` URL. This menu action calls:

```javascript
setFeedbackWebAppUrl('DEPLOYMENT_URL');
```

9. Choose **Tester Automation → Feedback → Run Feedback Smoke Test** and copy the fresh test-form URL shown in the success alert.
10. Submit test feedback from a private/incognito browser that does not have spreadsheet access.
11. Confirm one Feedback row was added and the related Activity Log row was completed.
12. When code changes later, use **Deploy → Manage deployments → Edit**, select the new version, and deploy the update. Existing `/exec` links continue using that deployment URL.

The three URLs are different:

- **Spreadsheet URL:** opens the private datastore spreadsheet for operators.
- **Apps Script project URL:** opens the private code editor and deployment controls.
- **Web App deployment URL:** ends in `/exec` and is the only base URL used in tester feedback links.

### Visible Menu Results

The spreadsheet menu now wraps initialization, health checks, all smoke tests, task seeding and assignment, reminder runs, and trigger installation/removal. Successful operations display a toast with counts or status. Missing configuration and failures display an alert in Google Sheets. The Feedback submenu can check whether a Web App URL is configured, run the smoke test, or show concise deployment instructions.

## Smart Tester Monitoring

Iteration 6 derives tester status from genuine participation recorded in Activity Log and Feedback. Reminder delivery is not participation. The monitoring engine uses calendar dates in the Apps Script project timezone, not elapsed 24-hour durations.

### Participation Summary

The following helpers expose structured monitoring data without changing tester status:

```javascript
getTesterParticipationSummary('TESTER_ID');
getTesterLastActivityDate('TESTER_ID');
getTesterCompletedActivityCount('TESTER_ID');
getTesterFeedbackCount('TESTER_ID');
determineTesterStatus('TESTER_ID');
```

Meaningful participation consists only of:

- An Activity Log row with `Status = Completed` and a valid `Completed At` timestamp
- A Feedback row with a valid `Submitted At` timestamp

`getTesterParticipationSummary()` returns current test day, assigned and completed activity counts, feedback count, last participation timestamp, days since participation, and the amount of participation recorded inside the tester's 14-day window. When no participation exists, inactivity is measured from the Start Date so a new tester receives a short grace period.

### Automatic Status Rules

Status precedence and recommendations are:

- `Inactive`: always preserved as a manual exclusion or opt-out. Monitoring never overwrites it.
- `Not Started`: no Start Date, or the Start Date is in the future.
- `Active`: currently within Day 1–14 and participation occurred today or one calendar day ago. A tester with no participation also remains Active during Day 1 and Day 2 as the initial grace period.
- `Needs Reminder`: currently within Day 1–14 with exactly two calendar days since meaningful participation, or Day 3 with no participation since starting.
- `At Risk`: three or more calendar days without meaningful participation. A tester whose 14-day period ends with no in-period participation is also At Risk.
- `Completed`: current day is after Day 14 and at least one completed activity or feedback submission occurred within that tester's Day 1–14 window.

Passing Day 14 alone never causes `Completed`. Automatic monitoring does not set `Inactive` and does not send special escalation emails.

`determineTesterStatus(testerId)` returns the recommendation and reason without mutation. `refreshTesterStatus(testerId)` applies a changed recommendation through the existing tester update service. `refreshAllTesterStatuses()` processes every tester independently and returns processed, changed, unchanged, failed, change, and error summaries without tokens.

### Monitoring Sheet

`initializeSpreadsheet()` now creates the operational `Monitoring` sheet idempotently with these headers:

```text
Tester ID | Name | Email | Current Day | Status | Completed Activities | Feedback Count | Last Participation At | Days Since Participation | Updated At
```

`refreshMonitoringSheet()` clears and rebuilds only rows below the Monitoring header. It writes one derived row per tester and never modifies Testers, Activity Log, or Feedback source rows. This is an operational view, not the final dashboard.

### Monitoring Refresh

Use **Tester Automation → Monitoring → Refresh Monitoring**, or call:

```javascript
runMonitoringRefresh();
```

This first refreshes all automatic tester statuses, then rebuilds the Monitoring sheet. The spreadsheet menu displays processed, changed, Active, Needs Reminder, At Risk, Completed, and failed counts in a toast or alert.

### Monitoring Trigger

Use the Monitoring menu or these functions:

```javascript
installMonitoringTrigger();
removeMonitoringTrigger();
```

The installer creates one daily `runMonitoringRefresh` time-driven trigger at approximately 7:00 PM in the project timezone. It is duplicate-safe. Removal deletes only triggers whose handler is `runMonitoringRefresh`; the 9:00 AM and 6:00 PM reminder triggers are untouched.

### Reminder Eligibility

Normal AM/PM task reminders now continue for `Active`, `Needs Reminder`, and `At Risk` testers. They remain disabled for `Not Started`, `Completed`, and manually `Inactive` testers. This change does not add escalation templates or change reminder content.

### Monitoring Smoke Test

Run **Tester Automation → Monitoring → Run Monitoring Smoke Test** or execute:

```javascript
runMonitoringSmokeTest();
```

The test creates isolated `[Smoke Test]` testers for future start, recent participation, two-day inactivity, three-day inactivity, post-Day-14 participation, and manual Inactive scenarios. It backdates only its own smoke-test activity to verify in-period completion, checks an idempotent second refresh, preserves Inactive, and confirms a Monitoring row is generated. Existing real tester source rows are not changed; persistent smoke-test rows remain for manual review.

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
