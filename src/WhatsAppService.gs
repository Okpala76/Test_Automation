function setEvolutionApiUrl(url) {
  assertWhatsAppOperator_();
  if (typeof url !== "string") {
    throw new Error("Evolution API URL is required.");
  }
  var normalized = url.trim().replace(/\/+$/, "");
  if (!/^https:\/\/[^\s]+$/i.test(normalized)) {
    throw new Error("Evolution API URL must be a valid HTTPS URL.");
  }
  PropertiesService.getScriptProperties().setProperty(
    EVOLUTION_API_URL_PROPERTY,
    normalized,
  );
}

function getEvolutionApiUrl() {
  assertWhatsAppOperator_();
  return getEvolutionApiUrl_();
}

function setEvolutionApiKey(apiKey) {
  assertWhatsAppOperator_();
  if (typeof apiKey !== "string" || apiKey.trim() === "") {
    throw new Error("Evolution API key is required.");
  }
  PropertiesService.getScriptProperties().setProperty(
    EVOLUTION_API_KEY_PROPERTY,
    apiKey.trim(),
  );
}

function getEvolutionApiKey() {
  assertWhatsAppOperator_();
  return getEvolutionApiKey_();
}

function setEvolutionInstance(instanceName) {
  assertWhatsAppOperator_();
  if (
    typeof instanceName !== "string" ||
    !/^[^\s/]+$/.test(instanceName.trim())
  ) {
    throw new Error(
      "Evolution instance name is required and cannot contain spaces or slashes.",
    );
  }
  PropertiesService.getScriptProperties().setProperty(
    EVOLUTION_INSTANCE_PROPERTY,
    instanceName.trim(),
  );
}

function getEvolutionInstance() {
  assertWhatsAppOperator_();
  return getEvolutionInstance_();
}

function setWhatsAppTestMode(enabled) {
  assertWhatsAppOperator_();
  PropertiesService.getScriptProperties().setProperty(
    WHATSAPP_TEST_MODE_PROPERTY,
    normalizeWhatsAppEnabled_(enabled) ? "true" : "false",
  );
}

function getWhatsAppTestMode() {
  assertWhatsAppOperator_();
  return getWhatsAppTestMode_();
}

function setWhatsAppTestRecipient(phone) {
  assertWhatsAppOperator_();
  var normalized = normalizeWhatsAppPhone(phone);
  if (!normalized) {
    throw new Error("A WhatsApp test recipient is required.");
  }
  PropertiesService.getScriptProperties().setProperty(
    WHATSAPP_TEST_RECIPIENT_PROPERTY,
    normalized,
  );
}

function getWhatsAppTestRecipient() {
  assertWhatsAppOperator_();
  return getWhatsAppTestRecipient_();
}

function getEvolutionApiUrl_() {
  return (
    PropertiesService.getScriptProperties().getProperty(
      EVOLUTION_API_URL_PROPERTY,
    ) || ""
  );
}

function getEvolutionApiKey_() {
  return (
    PropertiesService.getScriptProperties().getProperty(
      EVOLUTION_API_KEY_PROPERTY,
    ) || ""
  );
}

function getEvolutionInstance_() {
  return (
    PropertiesService.getScriptProperties().getProperty(
      EVOLUTION_INSTANCE_PROPERTY,
    ) || ""
  );
}

function getWhatsAppTestMode_() {
  return (
    PropertiesService.getScriptProperties().getProperty(
      WHATSAPP_TEST_MODE_PROPERTY,
    ) === "true"
  );
}

function getWhatsAppTestRecipient_() {
  return (
    PropertiesService.getScriptProperties().getProperty(
      WHATSAPP_TEST_RECIPIENT_PROPERTY,
    ) || ""
  );
}

function getWhatsAppConfigurationStatus() {
  assertWhatsAppOperator_();
  return getWhatsAppConfigurationStatus_();
}

function getWhatsAppConfigurationStatus_() {
  var testRecipient = getWhatsAppTestRecipient_();
  var testRecipientValid = false;
  if (testRecipient) {
    try {
      testRecipientValid =
        normalizeWhatsAppPhone(testRecipient) === testRecipient;
    } catch (error) {
      testRecipientValid = false;
    }
  }
  return {
    apiUrlConfigured: Boolean(getEvolutionApiUrl_()),
    apiKeyConfigured: Boolean(getEvolutionApiKey_()),
    instanceConfigured: Boolean(getEvolutionInstance_()),
    testMode: getWhatsAppTestMode_(),
    testRecipientConfigured: Boolean(testRecipient),
    testRecipientValid: testRecipientValid,
  };
}

function sendWhatsAppText(phone, text) {
  assertWhatsAppOperator_();
  return sendWhatsAppText_(phone, text);
}

function sendWhatsAppText_(phone, text) {
  var isTestMode = getWhatsAppTestMode_();
  var configuration;
  try {
    configuration = validateEvolutionConfiguration_();
  } catch (error) {
    return whatsappFailureResult_(error.message, 0, isTestMode);
  }
  if (typeof text !== "string" || text.trim() === "") {
    return whatsappFailureResult_(
      "A non-empty WhatsApp message is required.",
      0,
      isTestMode,
    );
  }

  var recipient;
  try {
    recipient = resolveWhatsAppRecipient_(phone);
  } catch (error) {
    return whatsappFailureResult_(error.message, 0, isTestMode);
  }

  var endpoint =
    configuration.apiUrl +
    "/message/sendText/" +
    encodeURIComponent(configuration.instance);
  var response;
  try {
    response = UrlFetchApp.fetch(endpoint, {
      method: "post",
      contentType: "application/json",
      headers: {
        apikey: configuration.apiKey,
      },
      payload: JSON.stringify({
        number: recipient,
        text: text,
      }),
      muteHttpExceptions: true,
    });
  } catch (error) {
    return whatsappFailureResult_(
      "Evolution API request failed because the server was unreachable or timed out.",
      0,
      isTestMode,
    );
  }

  var statusCode = response.getResponseCode();
  var responseData;
  try {
    responseData = JSON.parse(response.getContentText());
  } catch (error) {
    return whatsappFailureResult_(
      "Evolution API returned an invalid JSON response.",
      statusCode,
      isTestMode,
    );
  }

  if (statusCode < 200 || statusCode >= 300) {
    return whatsappFailureResult_(
      describeEvolutionFailure_(statusCode, responseData),
      statusCode,
      isTestMode,
    );
  }
  if (evolutionResponseIndicatesFailure_(responseData)) {
    return whatsappFailureResult_(
      describeEvolutionFailure_(statusCode, responseData),
      statusCode,
      isTestMode,
    );
  }
  if (!evolutionResponseConfirmsSend_(responseData)) {
    return whatsappFailureResult_(
      describeEvolutionFailure_(statusCode, responseData),
      statusCode,
      isTestMode,
    );
  }

  return {
    success: true,
    statusCode: statusCode,
    isTestMode: isTestMode,
    recipient: maskWhatsAppPhone_(recipient),
    messageId: getEvolutionMessageId_(responseData),
  };
}

function sendWhatsAppTaskReminder(tester, task, activity) {
  assertWhatsAppOperator_();
  return sendWhatsAppTaskReminder_(tester, task, activity);
}

function sendWhatsAppTaskReminder_(tester, task, activity) {
  if (!tester || !tester.testerId) {
    throw new Error("A valid tester is required for a WhatsApp reminder.");
  }
  if (!task || !task.taskId || !task.title) {
    throw new Error("A valid task is required for a WhatsApp reminder.");
  }
  if (!activity || !activity.activityId) {
    throw new Error("A valid activity is required for a WhatsApp reminder.");
  }

  return runWithReminderDeliveryLocks_(function () {
    var currentTester = getTesterById(tester.testerId);
    if (
      !currentTester ||
      REMINDER_ELIGIBLE_STATUSES.indexOf(currentTester.status) === -1
    ) {
      return {
        sent: false,
        reason: "not_eligible",
        activityId: activity.activityId,
      };
    }
    if (!currentTester.whatsappEnabled) {
      return {
        sent: false,
        reason: "whatsapp_disabled",
        activityId: activity.activityId,
      };
    }
    if (!currentTester.phone) {
      return {
        sent: false,
        reason: "missing_phone",
        activityId: activity.activityId,
      };
    }
    try {
      normalizeWhatsAppPhone(currentTester.phone);
    } catch (error) {
      return {
        sent: false,
        reason: "invalid_phone",
        activityId: activity.activityId,
      };
    }

    var currentTask = getTaskById(task.taskId);
    if (!currentTask || !currentTask.active) {
      return {
        sent: false,
        reason: "no_task",
        activityId: activity.activityId,
      };
    }
    var currentActivity = getActivityById(activity.activityId);
    if (
      !currentActivity ||
      currentActivity.testerId !== currentTester.testerId ||
      currentActivity.taskId !== currentTask.taskId
    ) {
      throw new Error(
        "The WhatsApp activity does not match the tester and task.",
      );
    }
    if (currentActivity.status === ACTIVITY_STATUSES.COMPLETED) {
      return {
        sent: false,
        reason: "already_completed",
        activityId: currentActivity.activityId,
      };
    }
    if (isAlreadyWhatsAppReminded_(currentActivity)) {
      return {
        sent: false,
        reason: "already_reminded",
        activityId: currentActivity.activityId,
      };
    }

    var result = sendWhatsAppText_(
      currentTester.phone,
      buildWhatsAppReminderMessage_(
        currentTester,
        currentTask,
        currentActivity,
      ),
    );
    if (!result.success) {
      throw new Error(result.error);
    }
    recordWhatsAppReminderSentWithoutLock_(currentActivity.activityId);
    return {
      sent: true,
      reason: "sent",
      activityId: currentActivity.activityId,
      taskId: currentTask.taskId,
      isTestMode: result.isTestMode,
      recipient: result.recipient,
      statusCode: result.statusCode,
    };
  });
}

function runWhatsAppSmokeTest() {
  try {
    assertWhatsAppOperator_();
    if (!getWhatsAppTestMode_()) {
      throw new Error(
        "WhatsApp Test Mode must be enabled before running the smoke test.",
      );
    }
    if (!getWhatsAppTestRecipient_()) {
      throw new Error(
        "A WhatsApp test recipient must be configured before the smoke test.",
      );
    }
    validateEvolutionConfiguration_();
    var result = sendWhatsAppText_(
      "",
      "Lodge Manager Tester Automation \u2705\n\nWhatsApp reminder integration is working.",
    );
    if (!result.success) {
      throw new Error(result.error);
    }
    showMenuToast_(
      "WhatsApp smoke test passed\nTest message sent to configured test recipient.",
      "Tester Automation",
    );
    return {
      ok: true,
      statusCode: result.statusCode,
      isTestMode: result.isTestMode,
      recipient: result.recipient,
    };
  } catch (error) {
    showMenuAlert_(
      "WhatsApp Smoke Test Failed",
      error && error.message
        ? error.message
        : "The WhatsApp test message could not be sent.",
    );
    return {
      ok: false,
      error:
        error && error.message
          ? error.message
          : "The WhatsApp test message could not be sent.",
    };
  }
}

function resolveWhatsAppRecipient_(testerPhone) {
  return resolveWhatsAppRecipientForMode_(
    testerPhone,
    getWhatsAppTestMode_(),
    getWhatsAppTestRecipient_(),
  );
}

function resolveWhatsAppRecipientForMode_(
  testerPhone,
  testMode,
  testRecipient,
) {
  if (testMode) {
    if (!testRecipient) {
      throw new Error(
        "WhatsApp Test Mode is enabled but no test recipient is configured. No message was sent.",
      );
    }
    return normalizeWhatsAppPhone(testRecipient);
  }
  var normalized = normalizeWhatsAppPhone(testerPhone);
  if (!normalized) {
    throw new Error("A valid WhatsApp phone is required.");
  }
  return normalized;
}

function validateEvolutionConfiguration_() {
  var apiUrl = getEvolutionApiUrl_();
  var apiKey = getEvolutionApiKey_();
  var instance = getEvolutionInstance_();
  var missing = [];
  if (!apiUrl) {
    missing.push("API URL");
  }
  if (!apiKey) {
    missing.push("API key");
  }
  if (!instance) {
    missing.push("instance");
  }
  if (missing.length > 0) {
    throw new Error(
      "Evolution API configuration is incomplete: " + missing.join(", ") + ".",
    );
  }
  if (!/^https:\/\/[^\s]+$/i.test(apiUrl)) {
    throw new Error("Evolution API URL must be a valid HTTPS URL.");
  }
  if (!/^[^\s/]+$/.test(instance)) {
    throw new Error("Evolution instance configuration is invalid.");
  }
  return {
    apiUrl: apiUrl,
    apiKey: apiKey,
    instance: instance,
  };
}

function buildWhatsAppReminderMessage_(tester, task, activity) {
  var periodName =
    task.period === TASK_PERIODS.AM ? "Morning Test" : "Evening Test";
  return [
    "Hi " + (tester.name || "there") + " \uD83D\uDC4B",
    "",
    "Oga LandLord - Day " + task.day + " " + periodName,
    "",
    "Today's activity:",
    task.instructions || task.title,
    "",
    "Complete the activity, then submit your feedback here:",
    "",
    getFeedbackUrlForActivity(activity.activityId),
    "",
    "Thanks for helping us test Oga Landlord.",
  ].join("\n");
}

function whatsappFailureResult_(message, statusCode, isTestMode) {
  return {
    success: false,
    statusCode: statusCode || 0,
    isTestMode: Boolean(isTestMode),
    error: message,
  };
}

function evolutionResponseIndicatesFailure_(responseData) {
  if (!responseData || typeof responseData !== "object") {
    return true;
  }
  if (responseData.error || responseData.success === false) {
    return true;
  }
  if (typeof responseData.status === "number" && responseData.status >= 400) {
    return true;
  }
  var status = String(responseData.status || "").toLowerCase();
  return ["error", "failed", "disconnected"].indexOf(status) !== -1;
}

function evolutionResponseConfirmsSend_(responseData) {
  return Boolean(getEvolutionMessageId_(responseData));
}

function describeEvolutionFailure_(statusCode, responseData) {
  var serialized = "";
  try {
    serialized = JSON.stringify(responseData).toLowerCase();
  } catch (error) {
    serialized = "";
  }
  if (
    serialized.indexOf("disconnect") !== -1 ||
    serialized.indexOf("connection") !== -1 ||
    serialized.indexOf('"state":"close') !== -1
  ) {
    return "Evolution API could not send because the WhatsApp instance is disconnected.";
  }
  if (
    serialized.indexOf("instance") !== -1 &&
    serialized.indexOf("not found") !== -1
  ) {
    return "Evolution API could not find the configured WhatsApp instance.";
  }
  if (statusCode === 401 || statusCode === 403) {
    return "Evolution API rejected the configured credentials.";
  }
  if (statusCode >= 200 && statusCode < 300) {
    return "Evolution API did not return a confirmed message ID.";
  }
  return (
    "Evolution API returned HTTP " +
    statusCode +
    " and did not confirm delivery."
  );
}

function getEvolutionMessageId_(responseData) {
  if (
    responseData &&
    responseData.key &&
    typeof responseData.key.id === "string"
  ) {
    return responseData.key.id;
  }
  return "";
}

function maskWhatsAppPhone_(phone) {
  var value = String(phone || "");
  if (value.length <= 4) {
    return "****";
  }
  return new Array(value.length - 3).join("*") + value.slice(-4);
}

function assertWhatsAppOperator_() {
  var activeEmail = String(
    Session.getActiveUser().getEmail() || "",
  ).toLowerCase();
  if (!activeEmail) {
    throw new Error("Spreadsheet operator authorization is required.");
  }
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var authorizedEmails = spreadsheet.getEditors().map(function (user) {
    return String(user.getEmail() || "").toLowerCase();
  });
  var owner = spreadsheet.getOwner();
  if (owner) {
    authorizedEmails.push(String(owner.getEmail() || "").toLowerCase());
  }
  if (authorizedEmails.indexOf(activeEmail) === -1) {
    throw new Error("Spreadsheet operator authorization is required.");
  }
}
