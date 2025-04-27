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
    .addItem('📧 Process Emails Now', 'runEmailTaskProcessingNow')
    .addItem('📧 Process Specific Email', 'promptForEmailSubject')
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
  showStatus('Starting rescheduling process...');
  
  Promise.resolve()
    .then(() => {
      // First delete existing calendar events
      const calendar = CalendarApp.getDefaultCalendar();
      const now = new Date();
      const oneMonthFromNow = new Date();
      oneMonthFromNow.setMonth(oneMonthFromNow.getMonth() + 1);
      
      const events = calendar.getEvents(now, oneMonthFromNow);
      events.forEach(event => {
        const desc = event.getDescription();
        if (desc && desc.includes('Priority: P')) {
          event.deleteEvent();
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
    });
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
  // --- Remove blatant execution marker ---
  // console.log(">>>> RUNNING LATEST Config.gs v1 - " + new Date().toISOString() + " <<<<");
  // --- End marker ---
  
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

        if (schedulingResult.success && schedulingResult.message !== 'Follow-up task skipped' && schedulingResult.message !== 'Paused task skipped') {
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
            
            // --- Set status to Pending for skipped Follow-ups if needed ---
            if (includesFollowUp) { // Use the calculated boolean
                console.log(` -> Skipped Follow-up task ${task.name}. Ensuring status is Pending.`);
                sheetManager.updateTaskStatus(task.row, 'Pending'); 
            }
            // --- End Status Update for Skipped Follow-ups ---

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
  
  let subject = '';
  let body = '';
  const taskDetails = `Task: ${task.name}\\nPriority: ${task.priority || 'Not set'}\\nTime Block: ${task.timeBlock || 'Default'} minutes\\nSource Task Row: ${task.row}\\nNotes: ${task.notes || 'None'}`; // Escaped newlines

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

    body = `Your task has been scheduled:\\n\\n${taskDetails}\\n\\nScheduled for: ${fullDateTimeStr}\\n\\nYou can view this task in your calendar.`;
  
  // Case 2: Task processing was successful, but it was skipped (Follow-up/Paused)
  } else if (schedulingResult && schedulingResult.success && schedulingResult.message && schedulingResult.message.toLowerCase().includes('skipped')) {
      subject = `Task Processed (Skipped): ${task.name}`;
      console.log(`CONFIRMATION EMAIL: Skipped task subject: "${subject}"`);
      body = `Your task was processed but skipped (not scheduled):\\n\\n${taskDetails}\\n\\nReason: ${schedulingResult.message}`;
      
  // Case 3: Task processing failed
  } else if (schedulingResult && !schedulingResult.success) {
      subject = `Error Processing Task: ${task.name}`;
      console.log(`CONFIRMATION EMAIL: Error processing task subject: "${subject}"`);
      body = `There was an error processing your task:\\n\\n${taskDetails}\\n\\nError: ${schedulingResult.message}`;
      
  // Case 4: Fallback / Unexpected scenario (e.g., schedulingResult is null/undefined)
  } else {
      subject = `Task Update: ${task.name}`;
      console.log(`CONFIRMATION EMAIL: Fallback task subject: "${subject}"`);
      body = `Task processed, but scheduling status is unclear:\\n\\n${taskDetails}\\n\\nScheduling Result: ${JSON.stringify(schedulingResult)}`;
  }


  try {
    console.log(`CONFIRMATION EMAIL BODY (Final Check):\\n---\\n${body}\\n---`); // Escaped newlines for multi-line body

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body.replace(/\\\\n/g, '\\n') // Replace escaped newlines with actual newlines for sending
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

// ----------- EMAIL PROCESSING FUNCTIONS -----------

/**
 * Process emails based on subject line to create tasks
 */
function processEmailsToTasks() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { // Wait 30 seconds
    console.log('Email processing skipped, lock not acquired.');
    return;
  }

  try {
    console.log('Starting email processing...');
    const threads = GmailApp.search('is:unread label:ProcessMyEmails'); // Use label to filter
    console.log(`Found ${threads.length} unread threads with label ProcessMyEmails.`);
    
    threads.forEach(thread => {
      const messages = thread.getMessages();
      const message = messages[0]; // Process only the first message in the thread
      
      const subject = message.getSubject();
      console.log(`Processing email thread: "${subject}"`);
      
      // --- Logic to extract task details from subject ---
      // Example: "Task: Project Alpha [P2, 60m]" 
      //          "Follow-up: Call John"
      //          "Quick Task: Send report"
      
      let taskName = subject;
      let priority = 'P1'; // Default to P1 if not specified
      let timeBlock = 30;  // Default time block
      let isFollowUp = false;

      // Check for follow-up keyword
      if (subject.toLowerCase().startsWith('follow-up:')) {
        isFollowUp = true;
        taskName = subject.substring(10).trim();
        priority = 'Follow-up'; // Set priority directly
        console.log(`Detected Follow-up task: "${taskName}"`);
      } 
      // Check for specific task format with details
      else if (subject.toLowerCase().startsWith('task:')) {
         taskName = subject.substring(5).trim();
         
         // Extract details like [P2, 60m]
         const detailMatch = taskName.match(/\\s*\\[(.*?)\\]$/); // Match [...] at the end
         if (detailMatch && detailMatch[1]) {
           const details = detailMatch[1].split(',').map(s => s.trim());
           taskName = taskName.substring(0, detailMatch.index).trim(); // Remove details part from name
           
           details.forEach(detail => {
             if (detail.toUpperCase().match(/^P[1-3]$/)) {
               priority = detail.toUpperCase();
             } else if (detail.toLowerCase().endsWith('m')) {
               const minutes = parseInt(detail.slice(0, -1), 10);
               if (!isNaN(minutes)) {
                 timeBlock = minutes;
               }
             }
           });
           console.log(`Extracted Task: Name="${taskName}", Priority="${priority}", TimeBlock=${timeBlock}`);
         } else {
            console.log(`Simple Task Detected: Name="${taskName}", Defaulting Priority=P1, TimeBlock=30`);
         }
      }
      // Check for Quick Task keyword 
      else if (subject.toLowerCase().startsWith('quick task:')) {
         taskName = subject.substring(11).trim();
         priority = 'P1'; // Quick tasks are P1
         timeBlock = 15; // Default quick tasks to 15 mins
         console.log(`Detected Quick Task: Name="${taskName}", Priority=P1, TimeBlock=15`);
      }
       else {
         // Treat as default P1 task if no keyword matches
         console.log(`Default Task Detected: Name="${taskName}", Defaulting Priority=P1, TimeBlock=30`);
       }

      // --- Add task to sheet ---
      try {
        const taskData = {
          name: taskName,
          priority: priority,
          timeBlock: timeBlock,
          notes: `From email: "${subject}"` 
        };
        
        // Add task and get row number
        const result = sheetManager.addTask(taskData); 
        
        // If it's a follow-up task, explicitly set status to Follow-up
        if (isFollowUp) {
          sheetManager.updateTaskStatus(result.row, 'Follow-up');
          console.log(`Set status to Follow-up for task in row ${result.row}`);
        }

        // Mark email as read and remove label
        message.markRead();
        thread.removeLabel(GmailApp.getUserLabelByName('ProcessMyEmails')); 
        console.log(`Successfully processed email "${subject}" into task row ${result.row}`);

      } catch (sheetError) {
         console.error(`Error adding task "${taskName}" from email to sheet: ${sheetError}`);
         // Optionally leave email unread or apply an error label
      }
      
    }); // End forEach thread
    
    console.log('Finished email processing.');
  } catch (error) {
    console.error('Error during email processing:', error);
  } finally {
    lock.releaseLock();
  }
}


/**
 * Creates a time-driven trigger for email processing if one doesn't exist.
 */
function setupEmailProcessingTrigger() {
  const functionName = 'processEmailsToTasks';
  const triggers = ScriptApp.getProjectTriggers();
  
  const triggerExists = triggers.some(trigger => 
    trigger.getHandlerFunction() === functionName
  );

  if (!triggerExists) {
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .everyMinutes(15) // Run every 15 minutes
      .create();
    console.log(`Created 15-minute trigger for ${functionName}.`);
    showStatus('Email processing trigger (15 min) created.');
  } else {
    console.log(`Trigger for ${functionName} already exists.`);
    showStatus('Email processing trigger already exists.');
  }
}


/**
 * Manually run email processing now (e.g., from menu)
 */
function runEmailTaskProcessingNow() {
  showStatus('Starting email processing...');
  processEmailsToTasks();
  showStatus('Email processing finished.');
}

/**
 * Prompt user for email subject to process manually
 */
function promptForEmailSubject() {
   const ui = SpreadsheetApp.getUi();
   const response = ui.prompt(
     'Process Specific Email', 
     'Enter the subject of the unread email to process:', 
     ui.ButtonSet.OK_CANCEL
   );

   if (response.getSelectedButton() == ui.Button.OK) {
     const subjectQuery = response.getResponseText();
     if (subjectQuery) {
       processSpecificEmail(subjectQuery.trim());
     } else {
       ui.alert('No subject entered.');
     }
   }
}


/**
 * Process a specific email by subject (should be unread)
 * @param {string} subject - The subject line to search for
 */
function processSpecificEmail(subject) {
  // Simplified version - searches for the first unread thread matching the subject
  try {
    console.log(`Searching for unread email with subject: "${subject}"`);
    const threads = GmailApp.search(`is:unread subject:("${subject}")`, 0, 1); // Find 1 match

    if (threads.length > 0) {
       const thread = threads[0];
       const message = thread.getMessages()[0];
       console.log('Found matching email. Processing...');
       
       // --- (Reuse logic similar to processEmailsToTasks) ---
      let taskName = message.getSubject(); // Use actual subject
      let priority = 'P1'; 
      let timeBlock = 30;
      let isFollowUp = false;

      if (taskName.toLowerCase().startsWith('follow-up:')) {
        isFollowUp = true;
        taskName = taskName.substring(10).trim();
        priority = 'Follow-up'; 
      } 
      else if (taskName.toLowerCase().startsWith('task:')) {
         taskName = taskName.substring(5).trim();
         const detailMatch = taskName.match(/\\s*\\[(.*?)\\]$/);
         if (detailMatch && detailMatch[1]) {
           const details = detailMatch[1].split(',').map(s => s.trim());
           taskName = taskName.substring(0, detailMatch.index).trim();
           details.forEach(detail => {
             if (detail.toUpperCase().match(/^P[1-3]$/)) { priority = detail.toUpperCase(); } 
             else if (detail.toLowerCase().endsWith('m')) { 
               const minutes = parseInt(detail.slice(0, -1), 10);
               if (!isNaN(minutes)) { timeBlock = minutes; }
             }
           });
         }
      }
      else if (taskName.toLowerCase().startsWith('quick task:')) {
         taskName = taskName.substring(11).trim();
         priority = 'P1'; 
         timeBlock = 15;
      }

      const taskData = { name: taskName, priority: priority, timeBlock: timeBlock, notes: `From email: "${message.getSubject()}"` };
      const result = sheetManager.addTask(taskData); 
      if (isFollowUp) { sheetManager.updateTaskStatus(result.row, 'Follow-up'); }
      
      message.markRead(); 
      // Optionally remove label if applicable
      try { thread.removeLabel(GmailApp.getUserLabelByName('ProcessMyEmails')); } catch(e) {} 
       
      showStatus(`Processed email "${subject}" into task row ${result.row}`);
    } else {
      showStatus(`No unread email found with subject: "${subject}"`);
    }
  } catch (error) {
     console.error('Error processing specific email:', error);
     showStatus('Error processing email: ' + error.message);
  }
}