function onEditTrigger(e) {
  doCalculate(e);

  scheduleRefreshMonitoring();
}

function scheduleRefreshMonitoring() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastScheduledTime = scriptProperties.getProperty('lastScheduledTime');
  const now = new Date().getTime();

  // Check if a trigger was scheduled in the last minute
  if (lastScheduledTime) {
    const elapsedTime = now - parseInt(lastScheduledTime, 10);
    if (elapsedTime < 60000) { // 60000 milliseconds = 1 minute
      Logger.log('Trigger already scheduled within the last minute. Skipping scheduling.');
      return;
    }
  }

  // Delete any existing triggers for refreshMonitoring to prevent duplicates
  deleteExistingTriggers('refreshMonitoring');

  // Create a new time-based trigger to run refreshMonitoring after 1 minute
  ScriptApp.newTrigger('refreshMonitoring')
    .timeBased()
    .after(1 * 60 * 1000) // Delay in milliseconds (1 minute)
    .create();

  // Update the last scheduled time
  scriptProperties.setProperty('lastScheduledTime', now.toString());
}

function deleteExistingTriggers(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}