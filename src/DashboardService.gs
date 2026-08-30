function getDashboardMetrics() {
  var testers = getAllTesters();
  var activities = getAllDashboardActivities_();
  var feedbackRecords = getAllDashboardFeedback_();
  var statusCounts = countTesterStatuses_();
  var aggregate = calculateDashboardAggregateMetrics_(
    activities,
    feedbackRecords,
    new Date()
  );

  var progress = testers.map(function (tester) {
    var participation = getTesterParticipationSummary(tester.testerId);
    return {
      testerId: tester.testerId,
      name: tester.name,
      email: tester.email,
      currentDay: participation.currentDay,
      status: tester.status,
      completedActivities: participation.completedActivities,
      feedbackCount: participation.feedbackCount,
      lastParticipationAt: participation.lastParticipationAt,
      daysSinceParticipation: participation.daysSinceParticipation
    };
  });

  var attentionRequired = progress
    .filter(function (tester) {
      return (
        tester.status === TESTER_STATUSES.AT_RISK ||
        tester.status === TESTER_STATUSES.NEEDS_REMINDER
      );
    })
    .sort(sortDashboardAttention_);

  progress.sort(function (first, second) {
    return String(first.name).localeCompare(String(second.name));
  });

  return {
    lastUpdated: new Date(),
    totalTesters: testers.length,
    statusCounts: statusCounts,
    today: aggregate.today,
    overall: aggregate.overall,
    attentionRequired: attentionRequired,
    testerProgress: progress
  };
}

function refreshDashboard() {
  refreshMonitoringSheet();
  return refreshDashboardFromCurrentMonitoring_();
}

function refreshDashboardFromCurrentMonitoring_() {
  var metrics = getDashboardMetrics();
  renderDashboard_(metrics);

  return {
    lastUpdated: metrics.lastUpdated,
    totalTesters: metrics.totalTesters,
    statusCounts: metrics.statusCounts,
    today: metrics.today,
    overall: metrics.overall,
    attentionRequired: metrics.attentionRequired.length,
    testerProgress: metrics.testerProgress.length
  };
}

function runDashboardSmokeTest() {
  var sourceBefore = dashboardSourceFingerprint_();
  var metrics = getDashboardMetrics();
  var statusTotal = Object.keys(TESTER_STATUSES).reduce(function (total, key) {
    return total + (metrics.statusCounts[TESTER_STATUSES[key]] || 0);
  }, 0);
  assertDashboardSmokeTest_(
    statusTotal === metrics.totalTesters,
    'Tester status counts do not equal the total tester count.'
  );

  var activities = getAllDashboardActivities_();
  var feedbackRecords = getAllDashboardFeedback_();
  var expectedTodayAssigned = activities.filter(function (activity) {
    return isDashboardDateToday_(activity.assignedAt);
  }).length;
  var expectedTodayCompleted = activities.filter(function (activity) {
    return (
      activity.status === ACTIVITY_STATUSES.COMPLETED &&
      isDashboardDateToday_(activity.completedAt)
    );
  }).length;
  var expectedTodayPending = activities.filter(function (activity) {
    return (
      isDashboardDateToday_(activity.assignedAt) &&
      activity.status !== ACTIVITY_STATUSES.COMPLETED
    );
  }).length;
  var expectedTodayFeedback = feedbackRecords.filter(function (feedback) {
    return isDashboardDateToday_(feedback.submittedAt);
  }).length;
  var expectedCompletedActivities = activities.filter(function (activity) {
    return (
      activity.status === ACTIVITY_STATUSES.COMPLETED &&
      isValidDashboardDate_(activity.completedAt)
    );
  }).length;
  var expectedFeedback = feedbackRecords.length;
  var expectedBugs = feedbackRecords.filter(function (feedback) {
    return isDashboardTrue_(feedback.bugReported);
  }).length;
  var ratings = feedbackRecords
    .map(function (feedback) {
      return Number(feedback.rating);
    })
    .filter(function (rating) {
      return isFinite(rating) && rating >= 1 && rating <= 5;
    });
  var expectedAverage =
    ratings.length === 0
      ? null
      : ratings.reduce(function (total, rating) {
          return total + rating;
        }, 0) / ratings.length;
  var expectedAttention = getAllTesters().filter(function (tester) {
    return (
      tester.status === TESTER_STATUSES.AT_RISK ||
      tester.status === TESTER_STATUSES.NEEDS_REMINDER
    );
  }).length;

  assertDashboardSmokeTest_(
    metrics.today.assignedActivities === expectedTodayAssigned &&
      metrics.today.completedActivities === expectedTodayCompleted &&
      metrics.today.pendingActivities === expectedTodayPending &&
      metrics.today.feedbackSubmissions === expectedTodayFeedback,
    'Today activity metrics are incorrect.'
  );
  assertDashboardSmokeTest_(
    metrics.overall.completedActivities === expectedCompletedActivities,
    'Completed activity count is incorrect.'
  );
  assertDashboardSmokeTest_(
    metrics.overall.feedbackSubmissions === expectedFeedback,
    'Total feedback count is incorrect.'
  );
  assertDashboardSmokeTest_(
    metrics.overall.bugsReported === expectedBugs,
    'Bug count is incorrect.'
  );
  assertDashboardSmokeTest_(
    dashboardNumbersEqual_(metrics.overall.averageFeedbackRating, expectedAverage),
    'Average feedback rating is incorrect.'
  );

  var expectedCompletionRate =
    activities.length === 0
      ? 0
      : expectedCompletedActivities / activities.length;
  assertDashboardSmokeTest_(
    dashboardNumbersEqual_(metrics.overall.completionRate, expectedCompletionRate),
    'Completion rate is incorrect.'
  );
  assertDashboardSmokeTest_(
    metrics.attentionRequired.length === expectedAttention,
    'Attention-required tester count is incorrect.'
  );

  var firstRefresh = refreshDashboard();
  var secondRefresh = refreshDashboard();
  assertDashboardSmokeTest_(
    firstRefresh.totalTesters === secondRefresh.totalTesters &&
      firstRefresh.attentionRequired === secondRefresh.attentionRequired &&
      dashboardNumbersEqual_(
        firstRefresh.overall.completionRate,
        secondRefresh.overall.completionRate
      ),
    'Repeated dashboard refreshes are not stable.'
  );

  var dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  assertDashboardSmokeTest_(
    dashboardSheet && dashboardSheet.getRange('A1').getValue() === 'Lodge Manager Closed Test Dashboard',
    'Dashboard sheet did not render correctly.'
  );
  var layout = getDashboardLayout_(metrics.attentionRequired.length, metrics.testerProgress.length);
  var explanations = getDashboardExplanations_();
  assertDashboardSmokeTest_(
    dashboardSheet.getRange(layout.summaryHelperRow, 1).getValue() === explanations.summary &&
      dashboardSheet.getRange(layout.todayHelperRow, 1).getValue() === explanations.today &&
      dashboardSheet.getRange(layout.overallHelperRow, 1).getValue() === explanations.overall &&
      dashboardSheet.getRange(layout.attentionHelperRow, 1).getValue() === explanations.attention &&
      dashboardSheet.getRange(layout.progressHelperRow, 1).getValue() === explanations.progress,
    'Dashboard helper text did not render correctly.'
  );
  assertDashboardSmokeTest_(
    dashboardSheet.getRange(layout.legendStartRow, 1).getValue() === TESTER_STATUSES.ACTIVE &&
      dashboardSheet.getRange(layout.legendStartRow + 5, 1).getValue() === TESTER_STATUSES.INACTIVE,
    'Dashboard status guide did not render correctly.'
  );
  assertDashboardSmokeTest_(
    sourceBefore === dashboardSourceFingerprint_(),
    'Dashboard refresh changed source data.'
  );

  return {
    ok: true,
    totalTesters: metrics.totalTesters,
    statusesCounted: statusTotal,
    todayAssigned: metrics.today.assignedActivities,
    todayCompleted: metrics.today.completedActivities,
    feedbackSubmissions: metrics.overall.feedbackSubmissions,
    bugsReported: metrics.overall.bugsReported,
    attentionRequired: metrics.attentionRequired.length,
    idempotent: true,
    sourceDataUnchanged: true
  };
}

function renderDashboard_(metrics) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!sheet) {
    throw new Error('The Dashboard sheet is missing. Run initializeSpreadsheet() first.');
  }

  var layout = getDashboardLayout_(
    metrics.attentionRequired.length,
    metrics.testerProgress.length
  );
  var explanations = getDashboardExplanations_();
  var requiredRows = layout.legendStartRow + getDashboardStatusGuide_().length;
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < 8) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 8 - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, sheet.getMaxRows(), 8).breakApart();
  sheet.clear();
  sheet.setConditionalFormatRules([]);
  sheet.getCharts().forEach(function (chart) {
    sheet.removeChart(chart);
  });

  sheet.getRange('A1:H1').merge().setValue('Lodge Manager Closed Test Dashboard');
  sheet
    .getRange('A2:H2')
    .merge()
    .setValue(
      'Last Updated: ' +
        Utilities.formatDate(
          metrics.lastUpdated,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd HH:mm:ss z'
        )
    );

  writeDashboardSection_(sheet, layout.summaryTitleRow, 'SUMMARY');
  writeDashboardHelper_(sheet, layout.summaryHelperRow, explanations.summary);
  sheet.getRange(layout.summaryHeaderRow, 1, 1, 7).setValues([[
    'TOTAL TESTERS',
    'ACTIVE',
    'NEEDS REMINDER',
    'AT RISK',
    'COMPLETED',
    'NOT STARTED',
    'INACTIVE'
  ]]);
  sheet.getRange(layout.summaryValueRow, 1, 1, 7).setValues([[
    metrics.totalTesters,
    metrics.statusCounts[TESTER_STATUSES.ACTIVE] || 0,
    metrics.statusCounts[TESTER_STATUSES.NEEDS_REMINDER] || 0,
    metrics.statusCounts[TESTER_STATUSES.AT_RISK] || 0,
    metrics.statusCounts[TESTER_STATUSES.COMPLETED] || 0,
    metrics.statusCounts[TESTER_STATUSES.NOT_STARTED] || 0,
    metrics.statusCounts[TESTER_STATUSES.INACTIVE] || 0
  ]]);

  writeDashboardSection_(sheet, layout.todayTitleRow, 'TODAY');
  writeDashboardHelper_(sheet, layout.todayHelperRow, explanations.today);
  sheet.getRange(layout.todayHeaderRow, 1, 1, 4).setValues([[
    "TODAY'S ASSIGNED ACTIVITIES",
    "TODAY'S COMPLETED ACTIVITIES",
    "TODAY'S PENDING ACTIVITIES",
    "TODAY'S FEEDBACK SUBMISSIONS"
  ]]);
  sheet.getRange(layout.todayValueRow, 1, 1, 4).setValues([[
    metrics.today.assignedActivities,
    metrics.today.completedActivities,
    metrics.today.pendingActivities,
    metrics.today.feedbackSubmissions
  ]]);

  writeDashboardSection_(sheet, layout.overallTitleRow, 'OVERALL');
  writeDashboardHelper_(sheet, layout.overallHelperRow, explanations.overall);
  sheet.getRange(layout.overallHeaderRow, 1, 1, 6).setValues([[
    'TOTAL ACTIVITIES ASSIGNED',
    'TOTAL ACTIVITIES COMPLETED',
    'COMPLETION RATE',
    'TOTAL FEEDBACK SUBMISSIONS',
    'TOTAL BUGS REPORTED',
    'AVERAGE RATING (1-5)'
  ]]);
  sheet.getRange(layout.overallValueRow, 1, 1, 6).setValues([[
    metrics.overall.assignedActivities,
    metrics.overall.completedActivities,
    metrics.overall.completionRate,
    metrics.overall.feedbackSubmissions,
    metrics.overall.bugsReported,
    metrics.overall.averageFeedbackRating === null
      ? 'N/A'
      : metrics.overall.averageFeedbackRating
  ]]);
  sheet
    .getRange(layout.overallHeaderRow, 3)
    .setNote('Completed assigned activities divided by total assigned activities.');
  sheet
    .getRange(layout.overallHeaderRow, 6)
    .setNote('Average tester feedback score from 1-5.');
  sheet.getRange(layout.overallValueRow, 3).setNumberFormat('0.0%');
  if (metrics.overall.averageFeedbackRating !== null) {
    sheet.getRange(layout.overallValueRow, 6).setNumberFormat('0.00');
  }

  writeDashboardSection_(sheet, layout.attentionTitleRow, 'ATTENTION REQUIRED');
  writeDashboardHelper_(sheet, layout.attentionHelperRow, explanations.attention);
  sheet.getRange(layout.attentionHeaderRow, 1, 1, 8).setValues([[
    'Name',
    'Email',
    'Current Day',
    'Status',
    'Last Participation',
    'Days Since Participation',
    'Completed Activities',
    'Feedback Count'
  ]]);
  if (metrics.attentionRequired.length > 0) {
    sheet
      .getRange(layout.attentionStartRow, 1, metrics.attentionRequired.length, 8)
      .setValues(
        metrics.attentionRequired.map(function (tester) {
          return [
            tester.name,
            tester.email,
            tester.currentDay,
            tester.status,
            tester.lastParticipationAt || '',
            tester.daysSinceParticipation === null
              ? ''
              : tester.daysSinceParticipation,
            tester.completedActivities,
            tester.feedbackCount
          ];
        })
      );
    sheet
      .getRange(layout.attentionStartRow, 5, metrics.attentionRequired.length, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm');
  }

  writeDashboardSection_(sheet, layout.progressTitleRow, 'TESTER PROGRESS');
  writeDashboardHelper_(sheet, layout.progressHelperRow, explanations.progress);
  sheet.getRange(layout.progressHeaderRow, 1, 1, 6).setValues([[
    'Name',
    'Current Day',
    'Status',
    'Completed Activities',
    'Feedback Count',
    'Last Participation'
  ]]);
  if (metrics.testerProgress.length > 0) {
    sheet
      .getRange(layout.progressStartRow, 1, metrics.testerProgress.length, 6)
      .setValues(
        metrics.testerProgress.map(function (tester) {
          return [
            tester.name,
            tester.currentDay,
            tester.status,
            tester.completedActivities,
            tester.feedbackCount,
            tester.lastParticipationAt || ''
          ];
        })
      );
    sheet
      .getRange(layout.progressStartRow, 6, metrics.testerProgress.length, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm');
  }

  writeDashboardSection_(sheet, layout.legendTitleRow, 'STATUS GUIDE');
  getDashboardStatusGuide_().forEach(function (entry, index) {
    var row = layout.legendStartRow + index;
    sheet.getRange(row, 1).setValue(entry.status);
    sheet.getRange(row, 2, 1, 7).merge().setValue(entry.meaning);
  });

  formatDashboard_(sheet, metrics, layout);
}

function writeDashboardSection_(sheet, row, title) {
  sheet.getRange(row, 1, 1, 8).merge().setValue(title);
}

function writeDashboardHelper_(sheet, row, text) {
  sheet.getRange(row, 1, 1, 8).merge().setValue(text);
}

function formatDashboard_(sheet, metrics, layout) {
  var darkGreen = '#214d3b';
  var sectionGreen = '#d9ead3';
  var headerFill = '#edf3ef';
  var border = '#c8d2ca';

  sheet
    .getRange('A1:H1')
    .setBackground(darkGreen)
    .setFontColor('#ffffff')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);
  sheet
    .getRange('A2:H2')
    .setBackground('#f4f7f5')
    .setFontColor('#516158')
    .setHorizontalAlignment('center');

  [
    layout.summaryTitleRow,
    layout.todayTitleRow,
    layout.overallTitleRow,
    layout.attentionTitleRow,
    layout.progressTitleRow,
    layout.legendTitleRow
  ].forEach(function (row) {
    sheet
      .getRange(row, 1, 1, 8)
      .setBackground(sectionGreen)
      .setFontWeight('bold')
      .setFontColor('#17382b');
  });

  [
    layout.summaryHeaderRow,
    layout.todayHeaderRow,
    layout.overallHeaderRow,
    layout.attentionHeaderRow,
    layout.progressHeaderRow
  ].forEach(function (row) {
    sheet
      .getRange(row, 1, 1, 8)
      .setBackground(headerFill)
      .setFontWeight('bold')
      .setWrap(true)
      .setVerticalAlignment('middle');
  });
  [
    layout.summaryHelperRow,
    layout.todayHelperRow,
    layout.overallHelperRow,
    layout.attentionHelperRow,
    layout.progressHelperRow
  ].forEach(function (row) {
    sheet
      .getRange(row, 1, 1, 8)
      .setBackground('#f7f9f8')
      .setFontColor('#5f6f66')
      .setFontStyle('italic')
      .setWrap(true);
  });
  sheet.setRowHeight(layout.overallHelperRow, 38);
  sheet.getRange(layout.summaryValueRow, 1, 1, 7).setFontSize(16).setFontWeight('bold');
  sheet.getRange(layout.todayValueRow, 1, 1, 4).setFontSize(15).setFontWeight('bold');
  sheet.getRange(layout.overallValueRow, 1, 1, 6).setFontSize(15).setFontWeight('bold');

  sheet.getRange(layout.summaryHeaderRow, 1, 2, 7).setBorder(
    true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(layout.todayHeaderRow, 1, 2, 4).setBorder(
    true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(layout.overallHeaderRow, 1, 2, 6).setBorder(
    true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet
    .getRange(
      layout.attentionHeaderRow,
      1,
      metrics.attentionRequired.length + 1,
      8
    )
    .setBorder(
      true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
    );
  sheet
    .getRange(
      layout.progressHeaderRow,
      1,
      metrics.testerProgress.length + 1,
      6
    )
    .setBorder(
      true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
    );
  sheet
    .getRange(layout.legendStartRow, 1, getDashboardStatusGuide_().length, 8)
    .setBorder(
      true, true, true, true, true, true, border, SpreadsheetApp.BorderStyle.SOLID
    );
  sheet
    .getRange(layout.legendStartRow, 1, getDashboardStatusGuide_().length, 1)
    .setFontWeight('bold');

  sheet.getRange(1, 1, sheet.getLastRow(), 8).setVerticalAlignment('middle');
  sheet.getRange(1, 1, sheet.getLastRow(), 8).setWrap(true);
  sheet.setFrozenRows(2);
  [190, 220, 120, 150, 165, 175, 150, 120].forEach(function (width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  applyDashboardConditionalFormatting_(
    sheet,
    metrics.attentionRequired.length,
    layout.attentionStartRow,
    metrics.testerProgress.length,
    layout.progressStartRow
  );
}

function getDashboardLayout_(attentionCount, progressCount) {
  var layout = {
    summaryTitleRow: 4,
    summaryHelperRow: 5,
    summaryHeaderRow: 6,
    summaryValueRow: 7,
    todayTitleRow: 9,
    todayHelperRow: 10,
    todayHeaderRow: 11,
    todayValueRow: 12,
    overallTitleRow: 14,
    overallHelperRow: 15,
    overallHeaderRow: 16,
    overallValueRow: 17,
    attentionTitleRow: 19,
    attentionHelperRow: 20,
    attentionHeaderRow: 21,
    attentionStartRow: 22
  };
  layout.progressTitleRow = layout.attentionStartRow + attentionCount + 1;
  layout.progressHelperRow = layout.progressTitleRow + 1;
  layout.progressHeaderRow = layout.progressTitleRow + 2;
  layout.progressStartRow = layout.progressTitleRow + 3;
  layout.legendTitleRow = layout.progressStartRow + progressCount + 1;
  layout.legendStartRow = layout.legendTitleRow + 1;
  return layout;
}

function getDashboardExplanations_() {
  return {
    summary: 'Current tester status based on recent activity and feedback.',
    today: 'Activity assigned, completed, and feedback submitted today.',
    overall: 'Performance across the full testing period. Completion Rate is completed assigned activities divided by total assigned activities. Average Rating is the average tester feedback score from 1-5.',
    attention: 'Testers who have gone 2+ days without meaningful participation.',
    progress: 'Current progress for every enrolled tester.'
  };
}

function getDashboardStatusGuide_() {
  return [
    { status: TESTER_STATUSES.ACTIVE, meaning: 'Participated recently' },
    { status: TESTER_STATUSES.NEEDS_REMINDER, meaning: 'No participation for 2 days' },
    { status: TESTER_STATUSES.AT_RISK, meaning: 'No participation for 3+ days' },
    { status: TESTER_STATUSES.COMPLETED, meaning: 'Finished testing period with participation' },
    { status: TESTER_STATUSES.NOT_STARTED, meaning: 'Testing has not begun' },
    { status: TESTER_STATUSES.INACTIVE, meaning: 'Manually excluded' }
  ];
}

function applyDashboardConditionalFormatting_(
  sheet,
  attentionCount,
  attentionStartRow,
  progressCount,
  progressStartRow
) {
  var statusStyles = {};
  statusStyles[TESTER_STATUSES.ACTIVE] = '#d9ead3';
  statusStyles[TESTER_STATUSES.NEEDS_REMINDER] = '#fff2cc';
  statusStyles[TESTER_STATUSES.AT_RISK] = '#f4cccc';
  statusStyles[TESTER_STATUSES.COMPLETED] = '#d9eaf7';
  statusStyles[TESTER_STATUSES.NOT_STARTED] = '#eeeeee';
  statusStyles[TESTER_STATUSES.INACTIVE] = '#e4d7ef';

  var ranges = [];
  if (attentionCount > 0) {
    ranges.push(sheet.getRange(attentionStartRow, 4, attentionCount, 1));
  }
  if (progressCount > 0) {
    ranges.push(sheet.getRange(progressStartRow, 3, progressCount, 1));
  }
  if (ranges.length === 0) {
    sheet.setConditionalFormatRules([]);
    return;
  }

  var rules = Object.keys(statusStyles).map(function (status) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(statusStyles[status])
      .setRanges(ranges)
      .build();
  });
  sheet.setConditionalFormatRules(rules);
}

function getAllDashboardFeedback_() {
  var context = getFeedbackSheetContext_();
  var lastRow = context.sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues()
    .filter(function (row) {
      return String(row[context.columnIndexes['Feedback ID'] - 1]).trim() !== '';
    })
    .map(function (row) {
      return feedbackObjectFromRow_(row, context.columnIndexes);
    });
}

function getAllDashboardActivities_() {
  return getAllActivities_();
}

function sortDashboardAttention_(first, second) {
  var firstPriority = first.status === TESTER_STATUSES.AT_RISK ? 0 : 1;
  var secondPriority = second.status === TESTER_STATUSES.AT_RISK ? 0 : 1;
  if (firstPriority !== secondPriority) {
    return firstPriority - secondPriority;
  }

  var firstDays = first.daysSinceParticipation === null ? -1 : first.daysSinceParticipation;
  var secondDays = second.daysSinceParticipation === null ? -1 : second.daysSinceParticipation;
  if (firstDays !== secondDays) {
    return secondDays - firstDays;
  }
  return String(first.name).localeCompare(String(second.name));
}

function isDashboardDateToday_(value) {
  return isDashboardSameCalendarDate_(value, new Date());
}

function isDashboardSameCalendarDate_(value, comparisonDate) {
  if (!isValidDashboardDate_(value) || !isValidDashboardDate_(comparisonDate)) {
    return false;
  }

  var timezone = Session.getScriptTimeZone();
  return (
    Utilities.formatDate(new Date(value), timezone, 'yyyy-MM-dd') ===
    Utilities.formatDate(new Date(comparisonDate), timezone, 'yyyy-MM-dd')
  );
}

function calculateDashboardAggregateMetrics_(activities, feedbackRecords, currentDate) {
  var todayAssigned = activities.filter(function (activity) {
    return isDashboardSameCalendarDate_(activity.assignedAt, currentDate);
  });
  var todayCompleted = activities.filter(function (activity) {
    return (
      activity.status === ACTIVITY_STATUSES.COMPLETED &&
      isDashboardSameCalendarDate_(activity.completedAt, currentDate)
    );
  });
  var todayPending = todayAssigned.filter(function (activity) {
    return activity.status !== ACTIVITY_STATUSES.COMPLETED;
  });
  var todayFeedback = feedbackRecords.filter(function (feedback) {
    return isDashboardSameCalendarDate_(feedback.submittedAt, currentDate);
  });
  var completedActivities = activities.filter(function (activity) {
    return (
      activity.status === ACTIVITY_STATUSES.COMPLETED &&
      isValidDashboardDate_(activity.completedAt)
    );
  });
  var validRatings = feedbackRecords
    .map(function (feedback) {
      return Number(feedback.rating);
    })
    .filter(function (rating) {
      return isFinite(rating) && rating >= 1 && rating <= 5;
    });

  return {
    today: {
      assignedActivities: todayAssigned.length,
      completedActivities: todayCompleted.length,
      pendingActivities: todayPending.length,
      feedbackSubmissions: todayFeedback.length
    },
    overall: {
      assignedActivities: activities.length,
      completedActivities: completedActivities.length,
      completionRate:
        activities.length === 0
          ? 0
          : completedActivities.length / activities.length,
      feedbackSubmissions: feedbackRecords.length,
      bugsReported: feedbackRecords.filter(function (feedback) {
        return isDashboardTrue_(feedback.bugReported);
      }).length,
      averageFeedbackRating:
        validRatings.length === 0
          ? null
          : validRatings.reduce(function (total, rating) {
              return total + rating;
            }, 0) / validRatings.length
    }
  };
}

function isValidDashboardDate_(value) {
  if (value === null || value === '' || typeof value === 'undefined') {
    return false;
  }
  return !isNaN(new Date(value).getTime());
}

function isDashboardTrue_(value) {
  var normalized = String(value).trim().toLowerCase();
  return value === true || normalized === 'true' || normalized === 'yes';
}

function dashboardNumbersEqual_(first, second) {
  if (first === null || second === null) {
    return first === second;
  }
  return Math.abs(first - second) < 0.0000001;
}

function dashboardSourceFingerprint_() {
  return JSON.stringify({
    testers: getAllTesters().map(function (tester) {
      return {
        testerId: tester.testerId,
        name: tester.name,
        email: tester.email,
        startDate: tester.startDate,
        status: tester.status,
        createdAt: tester.createdAt,
        updatedAt: tester.updatedAt
      };
    }),
    activities: getAllDashboardActivities_(),
    feedback: getAllDashboardFeedback_()
  });
}

function assertDashboardSmokeTest_(condition, message) {
  if (!condition) {
    throw new Error('Dashboard smoke test failed: ' + message);
  }
}
