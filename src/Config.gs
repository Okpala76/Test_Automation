var REQUIRED_SHEET_HEADERS = {
  'Testers': [
    'Tester ID',
    'Name',
    'Email',
    'Start Date',
    'Status',
    'Token',
    'Created At',
    'Updated At'
  ],
  'Tasks': [
    'Task ID',
    'Day',
    'Period',
    'Title',
    'Instructions',
    'Active'
  ],
  'Activity Log': [
    'Activity ID',
    'Tester ID',
    'Task ID',
    'Assigned At',
    'Completed At',
    'Status',
    'Last Reminder At',
    'Reminder Count'
  ],
  'Feedback': [
    'Feedback ID',
    'Tester ID',
    'Task ID',
    'Rating',
    'Completed',
    'Bug Reported',
    'Comment',
    'Submitted At'
  ],
  'Dashboard': [],
  'Monitoring': [
    'Tester ID',
    'Name',
    'Email',
    'Current Day',
    'Status',
    'Completed Activities',
    'Feedback Count',
    'Last Participation At',
    'Days Since Participation',
    'Updated At'
  ],
  'Guide': []
};

var TESTER_SHEET_NAME = 'Testers';

var TESTER_STATUSES = {
  NOT_STARTED: 'Not Started',
  ACTIVE: 'Active',
  NEEDS_REMINDER: 'Needs Reminder',
  AT_RISK: 'At Risk',
  COMPLETED: 'Completed',
  INACTIVE: 'Inactive'
};

var TESTER_EDITABLE_FIELDS = {
  name: 'Name',
  email: 'Email',
  startDate: 'Start Date',
  status: 'Status'
};

var TASK_SHEET_NAME = 'Tasks';
var ACTIVITY_LOG_SHEET_NAME = 'Activity Log';
var TESTING_PLAN_DAYS = 14;
var POST_TEST_DAY = 15;

var TASK_PERIODS = {
  AM: 'AM',
  PM: 'PM'
};

var ACTIVITY_STATUSES = {
  ASSIGNED: 'Assigned',
  COMPLETED: 'Completed',
  SKIPPED: 'Skipped'
};

var EMAIL_TEST_MODE_PROPERTY = 'EMAIL_TEST_MODE';
var TEST_EMAIL_RECIPIENT_PROPERTY = 'TEST_EMAIL_RECIPIENT';
var FEEDBACK_WEB_APP_URL_PROPERTY = 'FEEDBACK_WEB_APP_URL';
var FEEDBACK_SHEET_NAME = 'Feedback';
var FEEDBACK_COMMENT_MAX_LENGTH = 2000;

var REMINDER_TRIGGER_HANDLERS = {
  MORNING: 'sendMorningReminders',
  EVENING: 'sendEveningReminders'
};

var REMINDER_ELIGIBLE_STATUSES = [
  TESTER_STATUSES.ACTIVE,
  TESTER_STATUSES.NEEDS_REMINDER,
  TESTER_STATUSES.AT_RISK
];

var MONITORING_SHEET_NAME = 'Monitoring';
var MONITORING_TRIGGER_HANDLER = 'runMonitoringRefresh';
var GUIDE_SHEET_NAME = 'Guide';

var DEFAULT_TASKS = [
  {
    day: 1,
    period: 'AM',
    title: 'Open Lodge Manager and sign in',
    instructions: 'Open the app, sign in, and spend a few minutes becoming familiar with the first screen you see.'
  },
  {
    day: 1,
    period: 'PM',
    title: 'Reopen Lodge Manager',
    instructions: 'Reopen the app and confirm you can return to the main area without difficulty.'
  },
  {
    day: 2,
    period: 'AM',
    title: 'Inspect the tenant list',
    instructions: 'Open the tenant area and review the list and the information shown for a few tenants.'
  },
  {
    day: 2,
    period: 'PM',
    title: 'Revisit a tenant record',
    instructions: 'Open one tenant record again, then return to the tenant list using the app navigation.'
  },
  {
    day: 3,
    period: 'AM',
    title: 'Add a test tenant',
    instructions: 'Add a clearly labeled test tenant using safe test information only.'
  },
  {
    day: 3,
    period: 'PM',
    title: 'Confirm the test tenant was saved',
    instructions: 'Reopen the tenant area and confirm the test tenant you added is still present.'
  },
  {
    day: 4,
    period: 'AM',
    title: 'Review a tenancy record',
    instructions: 'Open a tenancy record and review the information presented there.'
  },
  {
    day: 4,
    period: 'PM',
    title: 'Revisit tenancy details',
    instructions: 'Return to the tenancy record and check that the details are easy to find again.'
  },
  {
    day: 5,
    period: 'AM',
    title: 'Inspect reminders',
    instructions: 'Open the reminders area and review the reminders and their displayed details.'
  },
  {
    day: 5,
    period: 'PM',
    title: 'Reopen reminders',
    instructions: 'Revisit the reminders area and notice whether anything about the navigation or information is unclear.'
  },
  {
    day: 6,
    period: 'AM',
    title: 'Edit test tenant information',
    instructions: 'Edit a non-sensitive field for the test tenant you created, then save the change.'
  },
  {
    day: 6,
    period: 'PM',
    title: 'Confirm the edited information',
    instructions: 'Reopen the test tenant record and confirm the edited information remains saved.'
  },
  {
    day: 7,
    period: 'AM',
    title: 'Navigate between major screens',
    instructions: 'Move between the main areas you have used so far, including tenants, tenancies, and reminders where available.'
  },
  {
    day: 7,
    period: 'PM',
    title: 'Try another navigation path',
    instructions: 'Use a different route through the app to return to an area you visited earlier.'
  },
  {
    day: 8,
    period: 'AM',
    title: 'Review tenant and tenancy information',
    instructions: 'Compare the information available from a tenant record and a tenancy record.'
  },
  {
    day: 8,
    period: 'PM',
    title: 'Revisit earlier test data',
    instructions: 'Return to the test tenant or tenancy information you used earlier and confirm it is still available.'
  },
  {
    day: 9,
    period: 'AM',
    title: 'Return after a break',
    instructions: 'Open the app after not using it for a while and navigate to a tenant or tenancy record.'
  },
  {
    day: 9,
    period: 'PM',
    title: 'Repeat a familiar task',
    instructions: 'Repeat one activity you completed earlier and note whether the steps still feel clear.'
  },
  {
    day: 10,
    period: 'AM',
    title: 'Inspect another tenant',
    instructions: 'Choose a different tenant from the list and review the available information.'
  },
  {
    day: 10,
    period: 'PM',
    title: 'Revisit a major screen',
    instructions: 'Return to a major screen you have not opened today and check that it is easy to reach.'
  },
  {
    day: 11,
    period: 'AM',
    title: 'Create or update safe test data',
    instructions: 'Use only test data to create or update one item in an area you have already explored.'
  },
  {
    day: 11,
    period: 'PM',
    title: 'Confirm saved test data',
    instructions: 'Reopen the item you created or updated and confirm the information remains visible.'
  },
  {
    day: 12,
    period: 'AM',
    title: 'Review reminders and tenant details',
    instructions: 'Review the reminders area and a tenant record during the same session.'
  },
  {
    day: 12,
    period: 'PM',
    title: 'Check a detail view again',
    instructions: 'Revisit one detailed record and pay attention to whether the displayed information is understandable.'
  },
  {
    day: 13,
    period: 'AM',
    title: 'Practice a typical app journey',
    instructions: 'Start from opening the app and navigate through the tenant or tenancy information you need.'
  },
  {
    day: 13,
    period: 'PM',
    title: 'Repeat the journey another way',
    instructions: 'Try a different navigation path to reach information you used in the morning.'
  },
  {
    day: 14,
    period: 'AM',
    title: 'Final beta app review',
    instructions: 'Open the app and revisit the key areas you used during the test.'
  },
  {
    day: 14,
    period: 'PM',
    title: 'Final usability revisit',
    instructions: 'Reopen the app, review anything that was confusing, and prepare to provide feedback later.'
  }
];
