/**
 * Gets work hours configuration
 * @returns {Object} Work hours configuration
 */
function getWorkHours() {
  const workHours = {
    'Monday': [
      { start: '11:00', end: '13:00' },
      { start: '14:00', end: '17:00' },
      { start: '20:00', end: '22:00' }
    ],
    'Tuesday': [
      { start: '11:00', end: '13:00' },
      { start: '14:00', end: '17:00' },
      { start: '20:00', end: '22:00' }
    ],
    'Wednesday': [
      { start: '11:00', end: '13:00' },
      { start: '14:00', end: '17:00' },
      { start: '20:00', end: '22:00' }
    ],
    'Thursday': [
      { start: '11:00', end: '13:00' },
      { start: '14:00', end: '17:00' },
      { start: '20:00', end: '22:00' }
    ],
    'Friday': [
      { start: '11:00', end: '13:00' },
      { start: '14:00', end: '17:00' },
      { start: '20:00', end: '22:00' }
    ]
  };
  return workHours;
}

/**
 * Initialize all settings
 * @returns {Object} All configuration settings
 */
function initializeSettings() {
  return {
    workHours: getWorkHours(),
    userEmail: PropertiesService.getScriptProperties().getProperty('userEmail'),
    timeZone: 'Asia/Kolkata',
    features: {
      optimizeShortTasks: true  // Set to false to disable
    }
  };
}

/**
 * Shows a status message to the user
 * @param {string} message - Message to show
 */
function showStatus(message) {
  SpreadsheetApp.getActive().toast(message, '🤖 Digital Assistant');
}

/**
 * Creates the menu in the spreadsheet - Simple Trigger Version
 * Note: This will run automatically on every refresh
 */
function onOpen(e) {
  // Create the simplified menu
  SpreadsheetApp.getUi()
    .createMenu('🤖 Digital Assistant')
    .addItem('📅 Reschedule All Tasks', 'startRescheduling') // Assumes startRescheduling function exists
    .addSeparator()
    .addItem('📧 Process Emails Now', 'processEmailTasks') // UPDATED
    .addItem('📧 Send Daily Digest Now', 'sendDailyTaskDigest')
    .addToUi();
}

/**
 * First-time setup for production use
 */
function installForProduction() {
  try {
    showStatus('Setting up system for production...');
    
    // Store user email in script properties
    const userEmail = Session.getActiveUser().getEmail();
    PropertiesService.getScriptProperties().setProperty('userEmail', userEmail);
    
    // Initialize task manager
    taskManager.initialize();
    
    // Initialize sheets
    sheetManager.initializeSheets();
    
    // Add Scheduled Time column if it doesn't exist
    addScheduledTimeColumn();
    
    // Create edit trigger
    const triggers = ScriptApp.getProjectTriggers();
    const hasEditTrigger = triggers.some(trigger => 
      trigger.getHandlerFunction() === 'onTaskEditTrigger'
    );
    
    if (!hasEditTrigger) {
      ScriptApp.newTrigger('onTaskEditTrigger')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
    }
    
    showStatus('System ready for production use!');
  } catch (error) {
    console.error('Installation error:', error);
    showStatus('Error during installation: ' + error.message);
  }
}

/**
 * Start rescheduling all tasks
 */
function startRescheduling() {
  // --- Add Lock Service ---
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { // Wait 10 seconds for the lock
    console.log('startRescheduling could not obtain lock. Another process likely running.');
    showStatus('Cannot start rescheduling now, another process is running. Please try again later.');
    return; 
  }
  console.log('startRescheduling acquired lock.');
  // --- End Lock Service ---

  showStatus('Starting rescheduling process...');
  
  Promise.resolve()
    .then(() => {
      // First delete existing calendar events
      const calendar = CalendarApp.getDefaultCalendar();
      const now = new Date();
      const oneMonthFromNow = new Date();
      oneMonthFromNow.setMonth(oneMonthFromNow.getMonth() + 1);
      
      const events = calendar.getEvents(now, oneMonthFromNow);
      console.log(`-> Found ${events.length} potentially overlapping events to check.`);

      events.forEach(event => {
        const eventStartTime = event.getStartTime(); // Get start time
        const eventEndTime = event.getEndTime();
        const eventTitle = event.getTitle();
        const eventId = event.getId();

        // --- Add Detailed Time Logging ---
        console.log(`--> Checking Event: "${eventTitle}"`);
        console.log(`    Event Start: ${eventStartTime.toISOString()} | Event End: ${eventEndTime.toISOString()}`);
        console.log(`    Check Window Start: ${now.toISOString()} | Check Window End (Now): ${now.toISOString()}`);
        // --- End Detailed Time Logging ---

        // Filter 1: Check if event ENDED within our time window (End is AFTER start check AND End is BEFORE or AT now)
        if (eventEndTime >= now && eventEndTime <= now) {
          // Filter 2: Check if it's one of our auto-scheduled events
          if (eventTitle && eventTitle.startsWith('Auto-Scheduled:')) {
            console.log(`  -> Event ENDED within check window and is Auto-Scheduled.`); // Log success
            // Filter 3: Check if we've already sent an email for this event
            const emailSentKey = `completionCheckSent_${eventId}`;
            if (!PropertiesService.getScriptProperties().getProperty(emailSentKey)) {
              console.log(`  -> Event "${eventTitle}" is Auto-Scheduled and not yet sent.`);
              event.deleteEvent();
            }
          }
        }
      });
      
      // Get all tasks and ensure they're pending
      const tasks = sheetManager.getTasks();
      tasks.forEach(task => {
        sheetManager.updateTaskStatus(task.row, 'Pending');
      });

      // Initialize task manager if needed
      if (!taskManager.settings) {
        taskManager.initialize();
      }
      
      // Process all tasks
      return taskManager.processPendingTasks();
    })
    .then(results => {
      console.log('Scheduled tasks:', results);
      
      // Count successful schedules
      const successfulSchedules = results.filter(result => result.success).length;
      
      showStatus(`Scheduled ${successfulSchedules} out of ${results.length} tasks. Check your calendar.`);
    })
    .catch(error => {
      console.error('Rescheduling failed:', error);
      showStatus('Error during rescheduling: ' + error.message);
    })
    .finally(() => { // --- Release Lock --- 
      lock.releaseLock();
      console.log('startRescheduling released lock.');
    }); // --- End Release Lock ---
}

/**
 * Clean up duplicate tasks
 */
function cleanupTasks() {
  try {
    showStatus('Cleaning up tasks...');
    
    const tasks = sheetManager.getTasks();
    const seen = new Set();
    const duplicates = [];
    
    tasks.forEach(task => {
      if (seen.has(task.name)) {
        duplicates.push(task.row);
      } else {
        seen.add(task.name);
      }
    });
    
    // Delete duplicate rows from bottom to top
    const tasksSheet = sheetManager.getTasksSheet();
    duplicates.sort((a, b) => b - a).forEach(row => {
      tasksSheet.deleteRow(row);
    });
    
    showStatus(`Cleaned up ${duplicates.length} duplicate tasks`);
  } catch (error) {
    console.error('Cleanup error:', error);
    showStatus('Error during cleanup: ' + error.message);
  }
}

/**
 * Reset everything to initial state
 */
function resetEverything() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset Everything',
    'This will delete all tasks. Are you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    showStatus('Reset cancelled');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActive();
    
    // Delete existing Tasks sheet
    const tasksSheet = ss.getSheetByName('Tasks');
    if (tasksSheet) ss.deleteSheet(tasksSheet);

    // Reinitialize sheets
    sheetManager.initializeSheets();
    
    showStatus('Reset completed successfully');
  } catch (error) {
    console.error('Reset error:', error);
    showStatus('Error during reset: ' + error.message);
  }
}

/**
 * Process the edited task after a delay
 */
function processPendingSheetTasks() {
  console.log(">>>> STARTING processPendingSheetTasks - LATEST VERSION CHECK <<<<");
  // --- Remove blatant execution marker ---
  // console.log(">>>> RUNNING LATEST Config.gs v1 - " + new Date().toISOString() + " <<<<");
  const lock = LockService.getScriptLock();
  // Allow more time as it might process multiple tasks
  if (!lock.tryLock(60000)) { // Wait up to 60 seconds
    console.log('Could not obtain lock for periodic task processing - another instance may be running. Skipping.');
    return; 
  }

  console.log('Periodic Check: Processing all pending tasks from sheet...');
  let processedCount = 0;
  let skippedCount = 0;

  try {
    // --- Get ALL pending tasks --- 
    const pendingTasks = sheetManager.getTasksByStatus('Pending');
    console.log(`Found ${pendingTasks.length} pending tasks.`);

    if (pendingTasks.length === 0) {
        console.log('No pending tasks found.');
        // Clear the lastEditedRow property in case it was left over (optional, but good hygiene)
        try { PropertiesService.getScriptProperties().deleteProperty('lastEditedRow'); } catch(e){} 
        return; // Exit if nothing to process
    }

    // --- Initialize Task Manager once --- 
    if (!taskManager.settings) {
      console.log('Initializing task manager');
      taskManager.initialize();
    }

    // --- Loop through pending tasks --- 
    for (const task of pendingTasks) {
        console.log(`Processing task: ${task.name} (Row: ${task.row}), Priority: ${task.priority || 'Not set'}`);
        const schedulingResult = taskManager.processTask(task); // processTask handles Follow-up/Pause check
        console.log(` -> Result: ${schedulingResult.message}`);

        if (schedulingResult.success && (!schedulingResult.message || !schedulingResult.message.toLowerCase().includes('skipped'))) {
            processedCount++;
            // Send confirmation email only if actually scheduled (processTask updates sheet status)
            const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
            if (userEmail) {
              try {
                // Pass the *original* task object and the scheduling result
                sendTaskCreationConfirmation(task, userEmail, schedulingResult);
              } catch (emailError) {
                console.error(`Error sending confirmation email for task ${task.name}: ${emailError}`);
              }
            } else {
              console.warn('User email not set, cannot send confirmation.');
            }
        } else if (!schedulingResult.success) {
            // Handle scheduling errors if needed - maybe update status to Error?
            console.error(`Failed to schedule task ${task.name}: ${schedulingResult.message}`);
            // --- Send Error Email ---
            const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
            if (userEmail) {
              sendTaskCreationConfirmation(task, userEmail, schedulingResult); // Send error notification
            }
            // --- End Send Error Email ---
        } else { // Implicitly: schedulingResult.success is true, but message indicates skipped
            skippedCount++; // Count skipped Follow-up/Paused
            
            // --- Remove diagnostic logging ---
            // console.log(`DEBUG: Checking skip message for status update. Message: "${schedulingResult.message}"`);
            const includesFollowUp = schedulingResult.message && schedulingResult.message.toLowerCase().includes('follow-up');
            // console.log(`DEBUG: Does message include 'follow-up'? ${includesFollowUp}`);
            // --- End diagnostic logging ---
            
            // --- Restore conditional status update ---
            // console.log(` -> Task "${task.name}" was skipped (Message: ${schedulingResult.message}). Forcing attempt to set status to Pending.`);
            // sheetManager.updateTaskStatus(task.row, 'Pending'); 
            // Original logic was conditional:
            if (includesFollowUp) { // Use the calculated boolean
                console.log(` -> Skipped Follow-up task ${task.name}. Setting status to Follow-up.`);
                sheetManager.updateTaskStatus(task.row, 'Follow-up'); // Set status to Follow-up AFTER skipping
            }
            // --- End restored conditional update ---

            // --- Send Skipped Email ---
            const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
            if (userEmail) {
              sendTaskCreationConfirmation(task, userEmail, schedulingResult); // Send skipped notification
            }
            // --- End Send Skipped Email ---
        }

        // --- Add post-processing status check ---
        try {
          const currentSheetStatus = sheetManager.getTaskStatusDirectly(task.row); // Requires new helper function
          console.log(`POST-PROCESSING CHECK: Task "${task.name}" (Row ${task.row}) final status in sheet: "${currentSheetStatus}"`);
        } catch (checkError) {
          console.error(`Error checking final status for task row ${task.row}: ${checkError}`);
        }
        // --- End post-processing status check ---

    }
    // --- End loop --- 

  } catch (error) {
    console.error('Error during periodic processing:', error);
  } finally {
    // --- Clear the lastEditedRow property (if it exists) --- 
    // This is less critical now but good hygiene 
    try {
        PropertiesService.getScriptProperties().deleteProperty('lastEditedRow');
        console.log(`Cleared lastEditedRow property (if it existed).`);
    } catch (deleteError) {
        console.error(`Error deleting lastEditedRow property: ${deleteError}`);
    }
    // ---------------------------------------------------------
    
    lock.releaseLock();
    console.log(`Periodic processing finished. Processed: ${processedCount}, Skipped: ${skippedCount}. Lock released.`);
  }
}

// Helper function to delete triggers calling a specific function
function deleteThisTrigger(functionName) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === functionName) {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }
    if (deletedCount > 0) {
        console.log(`Deleted ${deletedCount} trigger(s) for ${functionName}`);
    }
  } catch(e) {
      console.error(`Failed to delete triggers for ${functionName}: ${e}`);
  }
}

/**
 * Handles add-on installation
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Clean up and recreate all triggers
 */
function resetAllTriggers() {
  // Delete all existing triggers
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    console.log(`Deleting trigger: ${trigger.getHandlerFunction()}`);
    ScriptApp.deleteTrigger(trigger);
  });
  
  // Create onOpen trigger
  const openTrigger = ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  console.log('Created onOpen trigger:', openTrigger.getUniqueId());
    
  // Create edit trigger (only one)
  const editTrigger = ScriptApp.newTrigger('onTaskEditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  console.log('Created onTaskEdit trigger:', editTrigger.getUniqueId());
  
  showStatus('All triggers reset successfully');
}

/**
 * Check current triggers and log them
 */
function checkTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  console.log('=== Current Triggers ===');
  triggers.forEach(trigger => {
    console.log('Function:', trigger.getHandlerFunction());
    console.log('Event Type:', trigger.getEventType());
    console.log('Source:', trigger.getTriggerSource());
    console.log('---');
  });
}

/**
 * Add Scheduled Time column to Tasks sheet if it doesn't exist
 */
function addScheduledTimeColumn() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
    if (!sheet) {
      ui.alert('Error', 'Tasks sheet not found.', ui.ButtonSet.OK);
      return false;
    }
    
    // Get current headers
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Check if Scheduled Time column already exists
    const scheduledTimeIndex = headers.findIndex(header => 
      String(header).toLowerCase() === 'scheduled time'
    );
    
    if (scheduledTimeIndex >= 0) {
      console.log('Scheduled Time column already exists at index:', scheduledTimeIndex + 1);
      return true;
    }
    
    // Find the best position for the new column (after Status column)
    const statusIndex = headers.findIndex(header => 
      String(header).toLowerCase() === 'status'
    );
    
    const insertPosition = statusIndex >= 0 ? statusIndex + 2 : sheet.getLastColumn() + 1;
    
    // Insert the new column
    sheet.insertColumnAfter(insertPosition - 1);
    
    // Set the header
    sheet.getRange(1, insertPosition).setValue('Scheduled Time');
    
    // Format as date/time
    sheet.getRange(2, insertPosition, sheet.getLastRow() - 1, 1).setNumberFormat('M/d/yyyy h:mm');
    
    console.log('Added Scheduled Time column at position:', insertPosition);
    return true;
  } catch (error) {
    console.error('Error adding Scheduled Time column:', error);
    ui.alert('Error', 'Failed to add Scheduled Time column: ' + error.message, ui.ButtonSet.OK);
    return false;
  }
}

/**
 * Menu handler for standardizing follow-up format
 */
function standardizeFollowUpPrioritiesFromMenu() {
  standardizeFollowUpPriorities();
}

/**
 * Send confirmation email for task creation/processing
 * @param {Object} task - The task object (should include row)
 * @param {string} email - User email address
 * @param {Object} schedulingResult - The result object from processTask
 */
function sendTaskCreationConfirmation(task, email, schedulingResult) {
  console.log(`CONFIRMATION EMAIL: Preparing for task "${task.name}" with row=${task.row}`);
  console.log('CONFIRMATION EMAIL: Scheduling result:', JSON.stringify(schedulingResult));
  
  // --- Add Version Marker Log ---
  console.log('>>> Running sendTaskCreationConfirmation v3 (Config.gs) <<<'); 
  // --- End Version Marker Log ---

  let subject = '';
  let body = '';
  const taskDetails = `Task: ${task.name}\nPriority: ${task.priority || 'Not set'}\nTime Block: ${task.timeBlock || 'Default'} minutes\nSource Task Row: ${task.row}\nNotes: ${task.notes || 'None'}`;

  // Case 1: Task was successfully scheduled
  if (schedulingResult && schedulingResult.success && schedulingResult.events && schedulingResult.events.length > 0 && !schedulingResult.message.toLowerCase().includes('skipped')) { // Check for success AND not skipped
    const scheduledEvent = schedulingResult.events[0]; 
    const startTime = new Date(scheduledEvent.start); 

    if (!startTime || !(startTime instanceof Date) || isNaN(startTime)) {
       console.error(`CONFIRMATION EMAIL: Invalid start time received for task "${task.name}". Cannot format email.`);
       // Fallback to a simpler email? Or just return? Let's return for now.
       return;
    }
    console.log(`CONFIRMATION EMAIL: Event start time from result: ${startTime} (Type: ${typeof startTime})`);

    const timeZone = PropertiesService.getScriptProperties().getProperty('timeZone') || 'GMT';
    console.log(`CONFIRMATION EMAIL: Using time zone: ${timeZone}`);

    const timeStr = Utilities.formatDate(startTime, timeZone, '@h:mm a');
    const dateStr = Utilities.formatDate(startTime, timeZone, "'of' EEE (M/d)");
    const fullDateTimeStr = Utilities.formatDate(startTime, timeZone, 'EEEE, MMMM d, yyyy hh:mm a'); 
    console.log(`CONFIRMATION EMAIL: Formatted time: ${timeStr}, Formatted date: ${dateStr}, Full: ${fullDateTimeStr}`);

    subject = `Task Scheduled: ${task.name} ${timeStr} ${dateStr}`;
    console.log(`CONFIRMATION EMAIL: Scheduled task subject: "${subject}"`);

    body = `Your task has been scheduled:\n\n${taskDetails}\n\nScheduled for: ${fullDateTimeStr}\n\nYou can view this task in your calendar.`;
  
  // Case 2: Task processing was successful, but it was skipped (Follow-up/Paused)
  } else if (schedulingResult && schedulingResult.success && schedulingResult.message && schedulingResult.message.toLowerCase().includes('skipped')) {
      // --- Differentiate Follow-up skips from other skips (e.g., Paused) ---
      if (schedulingResult.message.toLowerCase().includes('follow-up')) {
          subject = `Follow-up Created: ${task.name}`; // Specific subject for Follow-ups
          console.log(`CONFIRMATION EMAIL: Follow-up Skipped task subject: "${subject}"`);
      } else {
          subject = `Task Processed (Skipped): ${task.name}`; // Generic subject for other skips
          console.log(`CONFIRMATION EMAIL: Generic Skipped task subject: "${subject}"`);
      }
      // --- End differentiation ---
      body = `Your task was processed but skipped (not scheduled):\n\n${taskDetails}\n\nReason: ${schedulingResult.message}`;
      
  // Case 3: Task processing failed
  } else if (schedulingResult && !schedulingResult.success) {
      subject = `Error Processing Task: ${task.name}`;
      console.log(`CONFIRMATION EMAIL: Error processing task subject: "${subject}"`);
      body = `There was an error processing your task:\n\n${taskDetails}\n\nError: ${schedulingResult.message}`;
      
  // Case 4: Fallback / Unexpected scenario (e.g., schedulingResult is null/undefined)
  } else {
      subject = `Task Update: ${task.name}`;
      console.log(`CONFIRMATION EMAIL: Fallback task subject: "${subject}"`);
      body = `Task processed, but scheduling status is unclear:\n\n${taskDetails}\n\nScheduling Result: ${JSON.stringify(schedulingResult)}`;
  }


  try {
    console.log(`CONFIRMATION EMAIL BODY (Final Check):\n---\n${body}\n---`); // Escaped newlines for multi-line body

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body
    });
    console.log(`CONFIRMATION EMAIL: Successfully sent email to ${email} for task "${task.name}"`);
  } catch (e) {
    console.error(`CONFIRMATION EMAIL: Failed to send email for task "${task.name}": ${e}`);
  }
}

/**
 * Test function for sending email
 */
function testEmailSending() {
  // This function is provided for testing purposes
  // It should be implemented to actually send an email
  console.log('Test email sending function called');
} 

// ----------- FOLLOW-UP LABEL PROCESSING ----------- 

/**
 * Processes emails labeled 'Process-Follow-Up' to create follow-up tasks.
 * Designed to run less frequently (e.g., every 6 hours).
 */
function processFollowUpLabels() {
  const lock = LockService.getScriptLock();
  // Wait up to 5 minutes for the lock
  if (!lock.tryLock(300000)) {
    console.log('Could not obtain lock for processFollowUpLabels - another instance may be running. Skipping.');
    return; 
  }

  const FOLLOW_UP_LABEL_NAME = 'Process-Follow-Up';
  const PROCESSED_LABEL_NAME = 'Task-Processed'; // Use the same processed label
  const MAX_THREADS_PER_RUN = 100; // Limit how many to process at once

  let processedCount = 0;
  let errorCount = 0;

  try {
    console.log(`====== STARTING Follow-Up Label Processing (${new Date().toLocaleString()}) ======`);
    
    // Ensure labels exist
    const followUpLabel = ensureLabelExists(FOLLOW_UP_LABEL_NAME);
    const processedLabel = ensureLabelExists(PROCESSED_LABEL_NAME);

    if (!followUpLabel) {
        console.error(`Critical Error: Could not find or create the label "${FOLLOW_UP_LABEL_NAME}". Aborting.`);
        return; // Cannot proceed without the label
    }
     if (!processedLabel) {
        console.error(`Critical Error: Could not find or create the label "${PROCESSED_LABEL_NAME}". Aborting.`);
        return; // Need the processed label too
    }

    // Search for threads with the follow-up label
    const threads = GmailApp.search(`label:${FOLLOW_UP_LABEL_NAME}`, 0, MAX_THREADS_PER_RUN);
    console.log(`Found ${threads.length} threads with label "${FOLLOW_UP_LABEL_NAME}".`);

    for (const thread of threads) {
      const threadId = thread.getId();
      console.log(`Processing Thread ID: ${threadId}`);

      try {
        const messages = thread.getMessages();
        if (!messages || messages.length === 0) {
          console.log(`Skipping thread ${threadId} - No messages found.`);
          // Remove label even if empty to avoid reprocessing
          thread.removeLabel(followUpLabel);
          continue;
        }

        // Use the first message for subject/details
        const message = messages[0]; 
        const subject = message.getSubject() || 'No Subject';
        const emailUrl = thread.getPermalink(); // Get permalink for the thread

        // Clean up task name
        let taskName = subject.replace(/^(Fwd|Re|FWD|RE):\s*/i, '').trim();
        taskName = `Follow up: ${taskName}`; // Prepend "Follow up:"

        console.log(` -> Task Name: "${taskName}"`);

        // Create task object
        const task = {
          name: taskName,
          priority: 'Follow-up', // Explicitly set priority
          timeBlock: 15,         // Default time for follow-ups (adjust if needed)
          notes: `Email Link: ${emailUrl}`,
          status: 'Follow-up'     // Set status directly to Follow-up
        };

        // Add task to sheet
        const addedRow = sheetManager.addTask(task);

        if (addedRow) {
          console.log(` -> Added task "${task.name}" to sheet row ${addedRow}. Status: Follow-up.`);
          processedCount++;
          
          // --- Post-processing ---
          // Remove the follow-up label
          thread.removeLabel(followUpLabel);
          // Add the standard processed label
          thread.addLabel(processedLabel);
          // Optionally archive
          if (thread.isInInbox()) {
            thread.moveToArchive();
            console.log(` -> Archived thread ${threadId}.`);
          }
          // ----------------------

        } else {
          console.error(` -> Failed to add task "${task.name}" (from thread ${threadId}) to sheet.`);
          errorCount++;
          // Consider adding an error label or leaving the Process-Follow-Up label for manual review
        }

      } catch (taskError) {
        console.error(` -> Error processing thread ${threadId}: ${taskError}`, taskError.stack);
        errorCount++;
        // Consider adding an error label or leaving the Process-Follow-Up label
      }
    } // End for loop

  } catch (error) {
    console.error('General Error during follow-up label processing:', error);
    errorCount++; // Count general errors too
  } finally {
    lock.releaseLock();
    console.log(`====== FINISHED Follow-Up Label Processing. Processed: ${processedCount}, Errors: ${errorCount}. Lock released. ======`);
  }
}

/**
 * Ensure a Gmail label exists. Returns the label object.
 * @param {string} labelName - Name of the label
 * @returns {GmailLabel|null} The GmailLabel object or null if creation failed.
 */
function ensureLabelExists(labelName) {
  console.log(`ENSURE_LABEL: Attempting to find label: "${labelName}"`);
  let label = null; // Initialize to null
  try {
      label = GmailApp.getUserLabelByName(labelName);
  } catch (e) {
      console.error(`ENSURE_LABEL: Error calling getUserLabelByName for "${labelName}": ${e}`);
      label = null; // Ensure label is null on error
  }
  
  console.log(`ENSURE_LABEL: Result of getUserLabelByName: ${label ? 'Found' : 'Not Found'}`);
  
  if (!label) {
    console.log(`ENSURE_LABEL: Label "${labelName}" not found, attempting to create.`);
    try {
      label = GmailApp.createLabel(labelName);
      console.log(`ENSURE_LABEL: Successfully created label: ${labelName}`);
    } catch (error) {
      console.error(`ENSURE_LABEL: Error creating label "${labelName}": ${error.message}`);
      // Check again in case of race condition or creation error
      console.log(`ENSURE_LABEL: Checking again after creation attempt...`);
      try {
           label = GmailApp.getUserLabelByName(labelName);
           console.log(`ENSURE_LABEL: Result of second getUserLabelByName: ${label ? 'Found' : 'Not Found'}`);
      } catch (e2) {
           console.error(`ENSURE_LABEL: Error on second getUserLabelByName for "${labelName}": ${e2}`);
           label = null;
      }
      if (!label) {
         console.error(`ENSURE_LABEL: Failed to ensure label "${labelName}" exists after creation attempt.`);
      }
    }
  } else {
    console.log(`ENSURE_LABEL: Label "${labelName}" already exists.`);
  }
  
  console.log(`ENSURE_LABEL: Returning label object for "${labelName}": ${label ? 'Exists' : 'NULL'}`);
  return label; // Return the label object or null
}

// ----------- WEB APP FUNCTIONS -----------

/**
 * Handles GET requests for the Web App.
 * This is triggered when a user clicks a link in the completion check email.
 * @param {Event} e - The event object containing request parameters.
 * @returns {HtmlOutput} An HTML page confirming the action.
 */
function doGet(e) {
  console.log('Web App doGet triggered.');
  console.log('Request Parameters:', JSON.stringify(e.parameter));

  const action = e.parameter.action;
  const rowStr = e.parameter.row;
  const user = Session.getActiveUser().getEmail();
  const scriptUser = PropertiesService.getScriptProperties().getProperty('userEmail');

  // --- Security Check: Ensure the person clicking is the script owner --- 
  if (!scriptUser || user !== scriptUser) {
    console.error('Web App Security Alert: Request received from unauthorized user:', user);
    return HtmlService.createHtmlOutput("<html><body><h1>Error</h1><p>You are not authorized to perform this action.</p></body></html>");
  }
  // --- End Security Check ---

  let message = 'Invalid request.'; // Default message
  let success = false;

  if (!action || !rowStr) {
    console.error('Web App Error: Missing action or row parameter.');
    message = 'Error: Missing required information in the request.';
  } else {
    const row = parseInt(rowStr, 10);
    if (isNaN(row)) {
      console.error('Web App Error: Invalid row number:', rowStr);
      message = 'Error: Invalid task identifier.';
    } else {
      console.log(`Web App Processing: Action='${action}', Row=${row}`);
      // --- Perform Action based on parameter --- 
      try {
        // Use a lock to prevent conflicts if multiple clicks happen quickly
        const lock = LockService.getScriptLock();
        if (lock.tryLock(5000)) { // Wait 5 seconds
          let taskName = '(Unknown Task)'; // Placeholder
          try {
             // Attempt to get task name for better messages (optional)
             const task = sheetManager.getTaskByRow(row);
             if (task) taskName = task.name;
          } catch (getNameError) { /* ignore error getting name */ }
          
          if (action === 'done') {
            success = sheetManager.deleteRow(row);
            if (success) message = `Task '${taskName}' (Row ${row}) marked as completed and removed.`;
            else message = `Failed to remove task '${taskName}' (Row ${row}). Check logs.`;
          } else if (action === 'reschedule') {
            success = sheetManager.updateTaskStatus(row, 'Pending');
            if (success) message = `Task '${taskName}' (Row ${row}) status set to Pending for rescheduling.`;
            else message = `Failed to update status for task '${taskName}' (Row ${row}). Check logs.`;
          } else if (action === 'asap') {
            success = sheetManager.updateTaskPriorityAndStatus(row, 'P1', 'Pending');
            if (success) message = `Task '${taskName}' (Row ${row}) status set to Pending and priority to P1 for ASAP rescheduling.`;
            else message = `Failed to update status/priority for task '${taskName}' (Row ${row}). Check logs.`;
          } else {
            message = `Error: Unknown action '${action}'.`;
          }
          lock.releaseLock();
        } else {
          message = 'Server busy, please try clicking the link again in a moment.';
          console.warn('Web App could not acquire lock.');
        }
      } catch (err) {
        console.error('Web App Error performing action:', err);
        message = 'An error occurred while processing your request. Please check the script logs.';
      }
      // --- End Action Logic --- 
    }
  }
  
  console.log('Web App Response Message:', message);
  // Return a simple HTML page to the user
  return HtmlService.createHtmlOutput(`<html><body style="font-family: sans-serif;"><h1>Digital Assistant Task Update</h1><p>${message}</p></body></html>`) 
    .setTitle('Task Update Confirmation');
}

function checkCompletedEvents() {
  console.log('Running Completion Check...');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { // Wait 30 seconds
    console.log('Completion Check skipped, lock not acquired.');
    return;
  }

  let webAppUrl = null;
  try {
    // --- Get Web App URL ---
    webAppUrl = ScriptApp.getService().getUrl();
    if (!webAppUrl) {
        console.error('Completion Check Error: Web App URL not available. Script must be deployed as a Web App.');
        lock.releaseLock(); // Release lock before returning
        return;
    }
    console.log('Web App URL:', webAppUrl);
    // --- End Get Web App URL ---

    const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
    if (!userEmail) {
      console.error('Completion Check Error: User email not set.');
      lock.releaseLock();
      return;
    }

    const calendar = CalendarApp.getDefaultCalendar();
    const scriptProperties = PropertiesService.getScriptProperties();
    const lastCheckTimestamp = scriptProperties.getProperty('lastCompletionCheckTimestamp');
    const now = new Date();
    const checkStartTime = lastCheckTimestamp ? new Date(parseInt(lastCheckTimestamp)) : new Date(now.getTime() - 60 * 60 * 1000); // Default to 1 hour ago if no previous check
    
    const maxLookbackMillis = 2 * 24 * 60 * 60 * 1000; // 2 days max lookback
    if (now.getTime() - checkStartTime.getTime() > maxLookbackMillis) {
        console.log(`Completion Check: Last check was too long ago (${checkStartTime.toISOString()}). Resetting check window to 1 hour ago.`);
        checkStartTime.setTime(now.getTime() - 60 * 60 * 1000);
    }

    console.log(`Completion Check: Searching for events ending between ${checkStartTime.toISOString()} and ${now.toISOString()}`);

    // Find potentially overlapping events
    const events = calendar.getEvents(checkStartTime, now); 
    let checkedCount = 0;
    console.log(`-> Found ${events.length} potentially overlapping events to check.`);

    events.forEach(event => {
      const eventStartTime = event.getStartTime(); // Get start time
      const eventEndTime = event.getEndTime();
      const eventTitle = event.getTitle();
      const eventId = event.getId();

      // --- Add Detailed Time Logging ---
      console.log(`--> Checking Event: "${eventTitle}"`);
      console.log(`    Event Start: ${eventStartTime.toISOString()} | Event End: ${eventEndTime.toISOString()}`);
      console.log(`    Check Window Start: ${checkStartTime.toISOString()} | Check Window End (Now): ${now.toISOString()}`);
      // --- End Detailed Time Logging ---

      // Filter 1: Check if event ENDED within our time window 
      if (eventEndTime >= checkStartTime && eventEndTime <= now) {
        console.log(`  -> Event "${eventTitle}" ENDED within check window.`); // Log filter 1 pass
        // Filter 2: Check if it's one of our auto-scheduled events
        if (eventTitle && eventTitle.startsWith('Auto-Scheduled:')) {
          console.log(`  -> Event is Auto-Scheduled.`); // Log filter 2 pass
          
          // Filter 3: Check if we've already sent an email for this event
          const emailSentKey = `completionCheckSent_${eventId}`;
          if (scriptProperties.getProperty(emailSentKey)) {
            console.log(`   -> Email already sent for event ID: ${eventId}. Skipping.`);
            return; // Skip this event
          }

          // Extract task row from description
          const desc = event.getDescription();
          const rowMatch = desc ? desc.match(/Source Task Row: (\d+)/) : null;
          if (!rowMatch || !rowMatch[1]) {
            console.warn(`   -> Could not extract row number from description for event: "${eventTitle}". Skipping.`);
            return; // Skip if no row number
          }
          const taskRow = parseInt(rowMatch[1], 10);

          // Get task status from sheet
          const taskStatus = sheetManager.getTaskStatusDirectly(taskRow);
          console.log(`   -> Task Row: ${taskRow}, Current Status in Sheet: "${taskStatus}"`);

          // Filter 4: Check if task status is still 'Scheduled'
          if (taskStatus === 'Scheduled') {
             console.log(`   -> Task status is 'Scheduled'. Proceeding to send email.`);
            const taskName = eventTitle.substring('Auto-Scheduled: '.length);
            const subject = `Task Completion Check: ${taskName} (Row ${taskRow})`; // Include row in subject

            // --- Construct Web App Links --- 
            const doneUrl = `${webAppUrl}?action=done&row=${taskRow}`;
            const rescheduleUrl = `${webAppUrl}?action=reschedule&row=${taskRow}`;
            const asapUrl = `${webAppUrl}?action=asap&row=${taskRow}`;
            // --- End Construct Links ---

            // --- Create HTML Body --- 
            const htmlBody = `
              <html>
                <body style="font-family: sans-serif;">
                  <p>Did you complete the task "<b>${taskName}</b>"?</p>
                  <p>
                    <a href="${doneUrl}" style="background-color: #4CAF50; color: white; padding: 10px 15px; text-align: center; text-decoration: none; display: inline-block; border-radius: 5px; margin-right: 10px;">Completed</a> 
                    <a href="${rescheduleUrl}" style="background-color: #ffcc00; color: black; padding: 10px 15px; text-align: center; text-decoration: none; display: inline-block; border-radius: 5px; margin-right: 10px;">No, Reschedule</a> 
                    <a href="${asapUrl}" style="background-color: #f44336; color: white; padding: 10px 15px; text-align: center; text-decoration: none; display: inline-block; border-radius: 5px;">No, Reschedule ASAP</a>
                  </p>
                  <hr>
                  <p style="font-size: smaller; color: grey;">
                    Task Details:<br>
                    Priority: ${desc ? (desc.match(/Priority: (.*)/) || [,''])[1] : 'N/A'}<br>
                    Time Block: ${desc ? (desc.match(/Time Block: (.*) minutes/) || [,''])[1] : 'N/A'} minutes<br>
                    Notes: ${desc ? (desc.match(/Notes: (.*)/) || [,''])[1] : 'N/A'}<br> 
                    (Row: ${taskRow}, Event ID: ${eventId})
                  </p>
                </body>
              </html>`;
            // --- End HTML Body --- 

            try {
              // Send HTML email
              MailApp.sendEmail({
                  to: userEmail, 
                  subject: subject, 
                  htmlBody: htmlBody // Use htmlBody instead of body
              });
              console.log(`    -> SENT completion check email for task "${taskName}" (Row ${taskRow})`);
              // Mark email as sent for this event ID
              scriptProperties.setProperty(emailSentKey, 'true'); 
              checkedCount++;
            } catch (e) {
              console.error(`   -> FAILED to send completion check email for task "${taskName}": ${e}`);
            }
          } else {
            console.log(`   -> Task status is "${taskStatus}", not "Scheduled". Skipping email.`);
          }
        } else {
            console.log(`  -> Event is not Auto-Scheduled. Skipping.`);
        }
      } else {
          console.log(`  -> Event "${eventTitle}" did NOT end within check window. Skipping.`);
      }
    }); // End forEach event

    // Update the timestamp for the next check
    scriptProperties.setProperty('lastCompletionCheckTimestamp', now.getTime().toString());
    console.log(`Completion Check finished. Sent ${checkedCount} emails. Updated last check timestamp.`);

  } catch (error) {
    console.error('Error during Completion Check:', error);
  } finally {
    // Ensure lock is released even if webAppUrl failed early
    if (lock && lock.hasLock()) {
        lock.releaseLock();
    }
  }
}

function specificallyFindEvent() {
  const eventTitleToFind = "Auto-Scheduled: AI Interviews - More rounds"; // Make sure this title is EXACTLY correct
  // Define a specific window around 5:45 PM IST (12:15 Z)
  const searchStart = new Date("2025-04-27T12:10:00Z"); // 5:40 PM IST
  const searchEnd = new Date("2025-04-27T12:20:00Z");   // 5:50 PM IST

  const calendar = CalendarApp.getDefaultCalendar();
  // Note: getEvents finds overlapping events, so we still filter by end time
  const events = calendar.getEvents(searchStart, searchEnd);
  Logger.log(`Specific Find: Found ${events.length} events overlapping ${searchStart.toISOString()} and ${searchEnd.toISOString()}`);

  let foundIt = false;
  events.forEach(event => {
    const eventEndStr = event.getEndTime().toISOString();
    const targetEndStr = "2025-04-27T12:15:00.000Z"; // Check this exact UTC time for 5:45 PM IST
    Logger.log(`--> Specific Find Checking: "${event.getTitle()}" | End: ${eventEndStr}`);
    // Use startsWith for title matching in case of slight variations, but exact match for time
    if (event.getTitle().startsWith(eventTitleToFind) && eventEndStr === targetEndStr) {
      Logger.log(`==> Specific Find MATCH FOUND: Event ID: ${event.getId()}`);
      foundIt = true;
    }
  });

  if (!foundIt) {
    Logger.log(`==> Specific Find: Did not find event "${eventTitleToFind}" ending exactly at ${targetEndStr} within the search window.`);
  }
  SpreadsheetApp.getUi().alert('Check Logs for specific event find results.');
}