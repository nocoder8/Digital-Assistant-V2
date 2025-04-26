/**
 * Email-to-Task Settings and Management
 */

/**
 * Enable the Email-to-Task feature
 */
function enableEmailTasks() {
  // Store setting
  PropertiesService.getScriptProperties().setProperty('emailTasksEnabled', 'true');
  
  // Create trigger
  createEmailCheckTrigger();
  
  // Confirm to user
  showStatus('Email-to-Task feature enabled. Checking for emails every 5 minutes.');
}

/**
 * Disable the Email-to-Task feature
 */
function disableEmailTasks() {
  // Store setting
  PropertiesService.getScriptProperties().setProperty('emailTasksEnabled', 'false');
  
  // Remove trigger
  deleteEmailCheckTrigger();
  
  // Confirm to user
  showStatus('Email-to-Task feature disabled.');
}

/**
 * Toggle auto-scheduling of email tasks
 * @param {boolean} enable - Whether to enable auto-scheduling
 */
function toggleEmailTaskAutoSchedule(enable) {
  PropertiesService.getScriptProperties().setProperty('emailTaskAutoSchedule', enable ? 'true' : 'false');
  showStatus(`Auto-scheduling for email tasks ${enable ? 'enabled' : 'disabled'}.`);
}

/**
 * Create trigger to check for task emails every 5 minutes
 */
function createEmailCheckTrigger() {
  // Delete existing triggers first
  deleteEmailCheckTrigger();
  
  // Create new trigger
  const trigger = ScriptApp.newTrigger('processEmailTasks')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  console.log('Created email check trigger:', trigger.getUniqueId());
}

/**
 * Delete email check triggers
 */
function deleteEmailCheckTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processEmailTasks') {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  
  if (count > 0) {
    console.log(`Deleted ${count} email check triggers`);
  }
}

/**
 * Check status of Email-to-Task feature
 * @returns {Object} Status information
 */
function checkEmailTaskStatus() {
  const props = PropertiesService.getScriptProperties();
  const enabled = props.getProperty('emailTasksEnabled') === 'true';
  const autoSchedule = props.getProperty('emailTaskAutoSchedule') === 'true';
  
  // Check if trigger exists
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(trigger => trigger.getHandlerFunction() === 'processEmailTasks');
  
  const status = {
    enabled: enabled,
    autoSchedule: autoSchedule,
    triggerExists: hasTrigger
  };
  
  // Show status to user
  showStatus(`Email-to-Task: ${enabled ? 'Enabled' : 'Disabled'}, Auto-schedule: ${autoSchedule ? 'Enabled' : 'Disabled'}`);
  
  return status;
}

/**
 * Run email task processing manually
 */
function runEmailTaskProcessingNow() {
  showStatus('Manually checking for email tasks...');
  processEmailTasks();
  showStatus('Email task check completed');
}

/**
 * Enable auto-scheduling for email tasks
 * Wrapper function for menu item
 */
function enableEmailTaskAutoSchedule() {
  toggleEmailTaskAutoSchedule(true);
}

/**
 * Disable auto-scheduling for email tasks
 * Wrapper function for menu item
 */
function disableEmailTaskAutoSchedule() {
  toggleEmailTaskAutoSchedule(false);
}

/**
 * Retry processing emails that previously failed
 */
function retryFailedEmails() {
  showStatus('Retrying failed emails...');
  
  // Get user email from settings
  const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
  if (!userEmail) {
    showStatus('User email not configured. Please run setup first.');
    return;
  }
  
  // Search for emails with error label
  const taskEmailSearch = `from:${userEmail} label:Task-Error-Processed (subject:"#td:" OR subject:"#FUP")`;
  const gmail = GmailApp.search(taskEmailSearch);
  
  if (gmail.length === 0) {
    showStatus('No failed emails to retry.');
    return;
  }
  
  // Remove error label from all threads
  const errorLabel = GmailApp.getUserLabelByName('Task-Error-Processed');
  if (errorLabel) {
    gmail.forEach(thread => {
      thread.removeLabel(errorLabel);
    });
  }
  
  showStatus(`Found ${gmail.length} failed emails. They will be processed in the next run.`);
  
  // Process them immediately if requested
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Process Now?',
    `Found ${gmail.length} failed emails. Process them now?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    processEmailTasks();
    showStatus('Retry processing completed.');
  }
}

/**
 * Prompt for email subject to process
 */
function promptForEmailSubject() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Process Specific Email',
    'Enter part of the email subject to search for:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const subjectText = result.getResponseText();
    if (subjectText) {
      showStatus(`Searching for emails with subject containing: ${subjectText}`);
      const success = processSpecificEmail(subjectText);
      showStatus(success ? 'Email processed successfully' : 'Failed to process email');
    }
  }
} 