/**
 * Email-to-Task Processor
 * Processes emails with specific subject prefixes to create tasks
 */

/**
 * Main function to process emails and create tasks
 */
function processEmailTasks() {
  const lock = LockService.getScriptLock();
  // Wait up to 30 seconds for the lock to become available.
  if (!lock.tryLock(30000)) {
    console.log('Could not obtain lock - another instance may be running. Skipping this execution.');
    return; // Exit if lock couldn't be obtained
  }

  try {
    console.log('============ EMAIL TASK PROCESSING START ============');
    console.log(`Process started at: ${new Date().toLocaleString()}`);
    
    // Declare startRowCount at function scope so it's accessible everywhere in the function
    let startRowCount = 0;
    
    try {
      // Check spreadsheet state at start
      startRowCount = SpreadsheetApp.getActive().getSheetByName('Tasks').getLastRow();
      console.log(`PROCESS START: Spreadsheet has ${startRowCount} rows initially`);
      
      // Check if feature is enabled
      const settings = PropertiesService.getScriptProperties().getProperty('emailTasksEnabled');
      if (settings !== 'true') {
        console.log('Email-to-Task feature is disabled. Skipping processing.');
        return;
      }
      
      console.log('Starting email task processing');
      
      // Get user email from settings
      const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
      if (!userEmail) {
        console.error('User email not configured. Please run setup first.');
        return;
      }
      
      // Define search query with more flexibility
      // Use a broader search that will find more emails, then filter them in code
      const taskEmailSearch = `from:${userEmail} newer_than:1d -label:Task-Processed -label:Task-Error-Processed`;
      
      console.log('Using search query:', taskEmailSearch);
      
      // Search for matching emails
      const gmail = GmailApp.search(taskEmailSearch);
      console.log(`Found ${gmail.length} email threads matching the basic criteria`);
      
      // Filter threads that actually have task markers in the subject
      const taskThreads = gmail.filter(thread => {
        const subject = thread.getFirstMessageSubject();
        return subject.toLowerCase().includes('#td') || subject.toLowerCase().includes('#fup');
      });
      
      console.log(`Found ${taskThreads.length} threads with task markers (#td or #fup)`);
      
      // Track processed tasks to avoid duplicates
      const processedTaskNames = new Set();
      
      // Process each thread
      for (const thread of taskThreads) {
        const messages = thread.getMessages();
        
        // Only process the most recent message in each thread
        const message = messages[messages.length - 1];
        const subject = message.getSubject();
        const body = message.getPlainBody();
        console.log('Processing email:', subject);
        
        try {
          let task;
          // Get the thread ID for a direct link to the email
          const threadId = thread.getId();
          const emailUrl = `https://mail.google.com/mail/u/0/#all/${threadId}`;
          
          // Parse task based on format
          if (subject.toLowerCase().includes('#fup')) {
            task = parseFollowUpFormat(subject, emailUrl);
          } else if (subject.toLowerCase().includes('#td')) {
            task = parseTaskFormat(subject, emailUrl, body);
          }
          
          if (task) {
            // Check if we've already processed a task with this name
            if (processedTaskNames.has(task.name)) {
              console.log(`Skipping duplicate task: ${task.name}`);
              
              // Still mark as processed to avoid future processing
              message.markRead();
              ensureLabelExists('Task-Processed');
              thread.addLabel(GmailApp.getUserLabelByName('Task-Processed'));
              
              continue;
            }
            
            // Add to processed set
            processedTaskNames.add(task.name);
            
            console.log('Successfully parsed task:', JSON.stringify(task));
            
            // Add task to sheet using the existing SheetManager
            const addedTask = sheetManager.addTask(task);
            console.log('Task added to sheet:', JSON.stringify(addedTask));
            
            // --- Add the row number to the task object ---
            if (addedTask && addedTask.success && addedTask.row) {
              task.row = addedTask.row;
              console.log(`Added row number ${task.row} to task object for scheduling`);
            } else {
              console.warn('Could not get row number when adding task, status update might fail.');
            }
            // ----------------------------------------------
            
            // Process the task immediately if auto-scheduling is enabled and it's not a follow-up
            const autoSchedule = PropertiesService.getScriptProperties().getProperty('emailTaskAutoSchedule');
            let schedulingResult = null; // Initialize result
            let mailSent = false; // Flag to prevent double emails
            
            // --- Check if it's a Follow-up OR if Auto-Schedule is off ---
            if (task.priority === 'Follow-up') {
              // Create a mock result for Follow-ups to trigger correct email
              schedulingResult = {
                success: true,
                message: 'Task skipped - Follow-up from email processing',
                events: []
              };
              console.log('EMAIL PROCESSING: Created mock scheduling result for follow-up task.');
              // Send confirmation immediately for Follow-ups
              sendTaskCreationConfirmation(task, userEmail, schedulingResult);
              mailSent = true;
            } 
            // --- If Auto-Schedule is ON and it's NOT a Follow-up, try scheduling ---
            else if (autoSchedule === 'true') { 
              console.log('Auto-scheduling task');
              console.log(`AUTO-SCHEDULE: Starting for task "${task.name}" at row ${task.row}`);
              
              // Initialize task manager if needed
              if (!taskManager.settings) {
                taskManager.initialize();
                console.log('AUTO-SCHEDULE: Initialized task manager');
              }
              
              // Only try to schedule if calendar access is available
              if (taskManager.hasCalendarAccess) {
                console.log(`AUTO-SCHEDULE: Calendar access available, calling processTask for "${task.name}"`);
                
                const beforeRowCount = SpreadsheetApp.getActive().getSheetByName('Tasks').getLastRow();
                console.log(`AUTO-SCHEDULE: Before processTask, spreadsheet has ${beforeRowCount} rows`);
                
                schedulingResult = taskManager.processTask(task, task.name);
                
                const afterRowCount = SpreadsheetApp.getActive().getSheetByName('Tasks').getLastRow();
                console.log(`AUTO-SCHEDULE: After processTask, spreadsheet has ${afterRowCount} rows (change: ${afterRowCount - beforeRowCount})`);
                
                console.log('Scheduling result:', JSON.stringify(schedulingResult));

                // Send confirmation based on scheduling result
                sendTaskCreationConfirmation(task, userEmail, schedulingResult);
                mailSent = true;
              } else {
                console.warn('Calendar access not available, task created but not scheduled');
                // Send basic confirmation if scheduling wasn't attempted
                sendTaskCreationConfirmation(task, userEmail, { success: false, message: 'Task created, auto-scheduling skipped (no calendar access)' });
                mailSent = true;
              }
            }
            
            // Send basic confirmation ONLY if not a Follow-up AND auto-schedule is OFF AND mail hasn't been sent
            if (task.priority !== 'Follow-up' && autoSchedule !== 'true' && !mailSent) {
                console.log('EMAIL PROCESSING: Sending basic Task Created confirmation (Auto-schedule OFF).');
                sendTaskCreationConfirmation(task, userEmail, { success: true, message: 'Task created, auto-scheduling is off.' }); 
            }
            // --- End Updated Confirmation Logic ---
            
            // Mark as read and add processed label
            message.markRead();
            ensureLabelExists('Task-Processed');
            thread.addLabel(GmailApp.getUserLabelByName('Task-Processed'));
            
            // Archive if not already archived
            if (!thread.isInInbox()) {
              console.log('Thread already archived, skipping archive step');
            } else {
              thread.moveToArchive();
            }
          }
        } catch (error) {
          console.error('Error processing email:', error.message);
          
          // Send error notification with example formats
          sendErrorNotification(subject, error.message, userEmail);
          
          // Mark as read and add error-processed label to prevent reprocessing
          message.markRead();
          ensureLabelExists('Task-Error-Processed');
          thread.addLabel(GmailApp.getUserLabelByName('Task-Error-Processed'));
          
          // Don't archive error emails so user can see them
        }
      }
    } catch (error) {
      console.error('Error processing emails:', error);
    }
    
    // Check spreadsheet state at end
    const endRowCount = SpreadsheetApp.getActive().getSheetByName('Tasks').getLastRow();
    console.log(`PROCESS END: Spreadsheet has ${endRowCount} rows at end (change: ${endRowCount - startRowCount})`);
    console.log('============ EMAIL TASK PROCESSING COMPLETE ============');
  } finally {
    // ALWAYS release the lock
    lock.releaseLock();
    console.log('Lock released.');
  }
}

/**
 * Parse follow-up format (#FUP)
 * @param {string} subject - Email subject
 * @param {string} emailUrl - URL to the email
 * @returns {Object} Task object
 */
function parseFollowUpFormat(subject, emailUrl) {
  // If #FUP is at start, use the rest of the subject
  let taskName;
  if (subject.toLowerCase().startsWith('#fup')) {
    taskName = subject.substring(4).trim();
  } else {
    // If #FUP is in the middle, use everything after it
    taskName = subject.substring(subject.toLowerCase().indexOf('#fup') + 4).trim();
  }
  
  // Clean up task name by removing any Fwd: or Re: prefixes
  taskName = taskName.replace(/^(Fwd|Re|FWD|RE):\s*/i, '').trim();
  
  console.log('Parsed FUP task name:', taskName);
  
  return {
    name: taskName,
    priority: 'Follow-up',
    timeBlock: 30, // Default time block for follow-ups
    notes: `Email Link: ${emailUrl}`,
    status: 'Pending'
  };
}

/**
 * Parse task format (#td: or #td)
 * @param {string} subject - Email subject
 * @param {string} emailUrl - URL to the email
 * @param {string} body - Email body
 * @returns {Object} Task object
 */
function parseTaskFormat(subject, emailUrl, body) {
  // Check if subject contains #td: or #td
  const tdIndex = subject.toLowerCase().indexOf('#td:');
  const tdNoColonIndex = subject.toLowerCase().indexOf('#td');
  
  // Use the format that appears in the subject
  const startIndex = tdIndex >= 0 ? tdIndex + 4 : tdNoColonIndex + 3;
  
  // Extract the part after #td: or #td
  const tdPart = subject.substring(startIndex).trim();
  console.log('Parsing TD format part:', tdPart);
  
  // Clean up by removing Fwd: and Re: prefixes
  const cleanSubject = tdPart.replace(/^(Fwd|Re|FWD|RE):\s*/i, '').trim();
  
  // Split by comma and trim each part
  const parts = cleanSubject.split(',').map(part => part.trim());
  console.log('Split parts:', parts);
  
  // Create task with defaults
  const task = {
    name: parts[0],
    priority: 'P1',  // Default priority
    timeBlock: 30,   // Default time block
    notes: `Email Link: ${emailUrl}`,
    status: 'Pending'
  };
  
  // Override defaults if provided
  if (parts.length > 1 && parts[1]) {
    // Parse priority
    let priority = parts[1].toUpperCase();
    if (!priority.match(/^(P[123]|FOLLOW-UP)$/i)) {
      // Try to normalize the priority
      if (priority.match(/^[123]$/)) {
        priority = 'P' + priority;
      } else if (priority.match(/^(FUP|FOLLOWUP)$/i)) {
        priority = 'Follow-Up';
      } else {
        // If invalid priority, keep default and add to notes
        task.notes += `\nInvalid priority provided: ${parts[1]}. Using default P1.`;
      }
    }
    task.priority = priority;
  }
  
  // Parse time block if provided
  if (parts.length > 2 && parts[2]) {
    const parsedTime = parseInt(parts[2]);
    if (!isNaN(parsedTime) && parsedTime > 0) {
      task.timeBlock = parsedTime;
    }
  }
  
  // Add notes if provided
  if (parts.length > 3 && parts[3]) {
    task.notes += `\nNotes: ${parts[3]}`;
  }
  
  // Add deadline if provided
  if (parts.length > 4 && parts[4]) {
    task.deadline = parseDateString(parts[4]);
  }
  
  return task;
}

/**
 * Send confirmation email for task creation
 * @param {Object} task - The created task
 * @param {string} userEmail - User's email address
 * @param {Object} schedulingResult - Result from scheduling (optional)
 */
function sendTaskCreationConfirmation(task, userEmail, schedulingResult) {
  console.log(`CONFIRMATION EMAIL: Preparing for task "${task.name}" with row=${task.row}`);
  console.log(`CONFIRMATION EMAIL: Scheduling result: ${JSON.stringify(schedulingResult || {})}`);
  
  // Generate a unique key for this task notification
  let notificationKey;
  if (schedulingResult && schedulingResult.events && schedulingResult.events.length > 0) {
    notificationKey = `email_sent_${task.name}_${schedulingResult.events[0].id}`;
  } else {
    notificationKey = `email_sent_${task.name}_${new Date().toISOString().split('T')[0]}`;
  }
  
  const emailAlreadySent = PropertiesService.getScriptProperties().getProperty(notificationKey);
  if (emailAlreadySent) {
    console.log(`CONFIRMATION EMAIL: Email already sent. Key: ${notificationKey}`);
    return; 
  }
  
  let subject;
  let bodyDetails = ''; // Store details for the body
  
  // Check if scheduling was successful and event data exists
  if (schedulingResult && schedulingResult.success && schedulingResult.events && schedulingResult.events.length > 0 && schedulingResult.events[0].start) {
    // Get the start time (which should be a Date object or UTC string)
    const eventStartRaw = schedulingResult.events[0].start;
    console.log(`CONFIRMATION EMAIL: Event start time from result: ${eventStartRaw} (Type: ${typeof eventStartRaw})`);
    
    // Ensure we have a Date object
    const eventStartDate = (typeof eventStartRaw === 'string') ? new Date(eventStartRaw) : eventStartRaw;
    
    // Get the CORRECT timezone from script properties
    let timeZone = PropertiesService.getScriptProperties().getProperty('timeZone');
    if (!timeZone) {
        console.error("CONFIRMATION EMAIL: Timezone not set in Script Properties! Using GMT as fallback.");
        timeZone = 'GMT'; // Fallback assignment is now allowed
    }
    console.log(`CONFIRMATION EMAIL: Using time zone: ${timeZone}`);

    // --- Simplified Time Formatting ---
    // Format time like @3:30 PM using the correct timezone
    const timeStr = Utilities.formatDate(eventStartDate, timeZone, '@h:mm a'); 
    // Format date like of Wed (4/9)
    const dateStr = Utilities.formatDate(eventStartDate, timeZone, "'of' EEE (M/d)");
    // Format full date/time for body
    const fullDateTimeStr = Utilities.formatDate(eventStartDate, timeZone, 'EEEE, MMMM dd, yyyy hh:mm a');
    // ----------------------------------

    console.log(`CONFIRMATION EMAIL: Formatted time: ${timeStr}, Formatted date: ${dateStr}, Full: ${fullDateTimeStr}`);
    
    // Construct the new subject line
    subject = `Task Scheduled: ${task.name} ${timeStr} ${dateStr}`; // Changed subject prefix
    console.log(`CONFIRMATION EMAIL: Scheduled task subject: "${subject}"`);

    // Add scheduling info to body details
    bodyDetails = `The task has been scheduled for ${fullDateTimeStr}.\n\nYou can view this task in your calendar.`;

  } else {
    // For tasks that weren't scheduled or if scheduling failed
    subject = `Task Created: ${task.name}`; // Keep original subject
    console.log(`CONFIRMATION EMAIL: Unscheduled task subject: "${subject}"`);
    bodyDetails = 'The task has been added to your task list but could not be automatically scheduled.';
  }
  
  const body = `
Your task has been ${schedulingResult && schedulingResult.success ? 'scheduled' : 'created'}:

Name: ${task.name}
Priority: ${task.priority}
Time Block: ${task.timeBlock} minutes
${task.deadline ? 'Deadline: ' + task.deadline : ''}
${task.notes ? '\nNotes:\n' + task.notes : ''}

${bodyDetails}
`;

  try {
    MailApp.sendEmail({
      to: userEmail,
      subject: subject,
      body: body
    });
    
    // Store that we've sent an email for this task/event
    PropertiesService.getScriptProperties().setProperty(notificationKey, 'sent_' + new Date().toISOString());
    
    console.log(`CONFIRMATION EMAIL: Successfully sent email to ${userEmail} for task "${task.name}"`);
  } catch (error) {
    console.error(`CONFIRMATION EMAIL: Error sending email: ${error.message}`);
  }
}

/**
 * Send error notification for failed task creation
 * @param {string} subject - Original email subject
 * @param {string} errorMessage - Error message
 * @param {string} userEmail - User's email address
 */
function sendErrorNotification(subject, errorMessage, userEmail) {
  const errorBody = `
There was an error creating your task:
${errorMessage}

Example formats:
1. For Follow-ups (forwarded emails):
   Subject: #FUP Fwd: [Original Email Subject]
   (Just add #FUP at the start when forwarding)

2. For New Tasks:
   Subject: #td: Weekly Review, P1, 60, Additional notes, tomorrow
   (Format: Task Name, Priority, Minutes (optional), Notes (optional), Deadline (optional))

Please try again using one of these formats.`;

  MailApp.sendEmail({
    to: userEmail,
    subject: 'Error Creating Task: ' + subject,
    body: errorBody
  });
}

/**
 * Helper function to parse date strings
 * @param {string} dateStr - Date string to parse
 * @returns {string} Formatted date or empty string
 */
function parseDateString(dateStr) {
  dateStr = dateStr.toLowerCase().trim();
  const today = new Date();
  
  switch (dateStr) {
    case 'today':
      return formatDate(today);
    case 'tomorrow':
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);
      return formatDate(tomorrow);
    case 'next week':
      const nextWeek = new Date(today);
      nextWeek.setDate(nextWeek.getDate() + 7);
      return formatDate(nextWeek);
    default:
      // Try parsing as date string
      const date = new Date(dateStr);
      if (!isNaN(date.getTime())) {
        return formatDate(date);
      }
      return '';
  }
}

/**
 * Format date as YYYY-MM-DD
 * @param {Date} date - Date to format
 * @returns {string} Formatted date
 */
function formatDate(date) {
  return Utilities.formatDate(date, 
    PropertiesService.getScriptProperties().getProperty('timeZone') || 'GMT', 
    'yyyy-MM-dd');
}

/**
 * Ensure a Gmail label exists
 * @param {string} labelName - Name of the label
 */
function ensureLabelExists(labelName) {
  try {
    // Try to get the label
    const label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      // Create if it doesn't exist
      GmailApp.createLabel(labelName);
    }
  } catch (error) {
    console.error(`Error ensuring label ${labelName} exists:`, error);
    // Create the label anyway
    try {
      GmailApp.createLabel(labelName);
    } catch (e) {
      // Ignore if it already exists
    }
  }
}

/**
 * Process a specific email by subject (for debugging)
 * @param {string} subjectText - Text to search for in the subject
 */
function processSpecificEmail(subjectText) {
  // Get user email from settings
  const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
  if (!userEmail) {
    console.error('User email not configured. Please run setup first.');
    return;
  }
  
  // Search for the specific email
  const searchQuery = `from:${userEmail} subject:"${subjectText}"`;
  console.log('Searching for:', searchQuery);
  
  const threads = GmailApp.search(searchQuery);
  console.log(`Found ${threads.length} matching threads`);
  
  if (threads.length === 0) {
    console.log('No matching emails found');
    return;
  }
  
  // Process the first matching thread
  const thread = threads[0];
  const messages = thread.getMessages();
  const message = messages[0];
  
  console.log('Processing email:', message.getSubject());
  
  try {
    let task;
    // Get the thread ID for a direct link to the email
    const threadId = thread.getId();
    const emailUrl = `https://mail.google.com/mail/u/0/#all/${threadId}`;
    
    // Parse task based on format
    if (message.getSubject().toLowerCase().includes('#fup')) {
      task = parseFollowUpFormat(message.getSubject(), emailUrl);
    } else if (message.getSubject().toLowerCase().includes('#td')) {
      task = parseTaskFormat(message.getSubject(), emailUrl, message.getPlainBody());
    }
    
    if (task) {
      console.log('Successfully parsed task:', JSON.stringify(task));
      
      // Add task to sheet using the existing SheetManager
      const addedTask = sheetManager.addTask(task);
      console.log('Task added to sheet:', JSON.stringify(addedTask));
      
      // Process the task immediately if auto-scheduling is enabled and it's not a follow-up
      const autoSchedule = PropertiesService.getScriptProperties().getProperty('emailTaskAutoSchedule');
      let schedulingResult = null;
      
      if (autoSchedule === 'true' && task.priority !== 'Follow-Up') {
        console.log('Auto-scheduling task');
        
        // Initialize task manager if needed
        if (!taskManager.settings) {
          taskManager.initialize();
        }
        
        // Only try to schedule if calendar access is available
        if (taskManager.hasCalendarAccess) {
          schedulingResult = taskManager.processTask(task, task.name);
          console.log('Scheduling result:', JSON.stringify(schedulingResult));
        } else {
          console.warn('Calendar access not available, task created but not scheduled');
        }
      }
      
      // Send confirmation with or without scheduling info
      sendTaskCreationConfirmation(task, userEmail, schedulingResult);
      
      console.log('Task processing completed successfully');
      return true;
    } else {
      console.log('Failed to parse task from email');
      return false;
    }
  } catch (error) {
    console.error('Error processing email:', error.message);
    return false;
  }
}

/**
 * Test function to find task emails without processing them
 */
function testFindTaskEmails() {
  // Get user email from settings
  const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
  if (!userEmail) {
    console.error('User email not configured. Please run setup first.');
    return;
  }
  
  // Try different search queries
  const searchQueries = [
    // Basic search
    `from:${userEmail} newer_than:1d`,
    
    // Current search
    `from:${userEmail} newer_than:1d -label:Task-Processed -label:Task-Error-Processed -subject:"Task Scheduled:" -subject:"Task Created:" -subject:"Error Creating Task:" (subject:"#td:" OR subject:"#FUP")`,
    
    // Simplified search
    `from:${userEmail} newer_than:1d subject:#td`,
    
    // Another variation
    `from:${userEmail} newer_than:1d subject:"#td"`,
    
    // Just looking for #FUP
    `from:${userEmail} newer_than:1d subject:#FUP`
  ];
  
  const results = {};
  
  // Try each search query
  for (const [index, query] of searchQueries.entries()) {
    console.log(`Testing search query ${index + 1}: ${query}`);
    
    const threads = GmailApp.search(query);
    results[`query${index + 1}`] = {
      query: query,
      count: threads.length,
      threads: threads.map(thread => {
        const firstMessage = thread.getMessages()[0];
        return {
          subject: firstMessage.getSubject(),
          date: firstMessage.getDate().toLocaleString(),
          id: thread.getId()
        };
      })
    };
  }
  
  // Log the results
  console.log('Email search test results:', JSON.stringify(results, null, 2));
  
  // Show a summary to the user
  const ui = SpreadsheetApp.getUi();
  let message = 'Email Search Test Results:\n\n';
  
  for (const [key, result] of Object.entries(results)) {
    message += `${key}: Found ${result.count} emails\n`;
    
    if (result.count > 0) {
      message += 'First few subjects:\n';
      result.threads.slice(0, 3).forEach(thread => {
        message += `- ${thread.subject}\n`;
      });
      message += '\n';
    }
  }
  
  ui.alert('Email Search Test', message, ui.ButtonSet.OK);
  
  return results;
}

/**
 * Reset email processing labels
 */
function resetEmailLabels() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset Email Labels',
    'This will remove all Task-Processed and Task-Error-Processed labels, allowing emails to be processed again. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // Get the labels
    const processedLabel = GmailApp.getUserLabelByName('Task-Processed');
    const errorLabel = GmailApp.getUserLabelByName('Task-Error-Processed');
    
    // Count threads with these labels
    let processedCount = 0;
    let errorCount = 0;
    
    // Remove processed label from all threads
    if (processedLabel) {
      const threads = GmailApp.search(`label:${processedLabel.getName()}`);
      processedCount = threads.length;
      
      threads.forEach(thread => {
        thread.removeLabel(processedLabel);
      });
    }
    
    // Remove error label from all threads
    if (errorLabel) {
      const threads = GmailApp.search(`label:${errorLabel.getName()}`);
      errorCount = threads.length;
      
      threads.forEach(thread => {
        thread.removeLabel(errorLabel);
      });
    }
    
    ui.alert(
      'Labels Reset',
      `Removed Task-Processed label from ${processedCount} threads\nRemoved Task-Error-Processed label from ${errorCount} threads`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    console.error('Error resetting labels:', error);
    ui.alert('Error', 'Failed to reset labels: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Clean up duplicate tasks in the spreadsheet
 */
function cleanupDuplicateTasks() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clean Up Duplicate Tasks',
    'This will remove duplicate tasks with the same name, keeping only the most recent one. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // Get all tasks
    const tasks = sheetManager.getTasks();
    
    // Track unique task names and their rows
    const uniqueTasks = new Map();
    const duplicates = [];
    
    // Find duplicates (keep the highest row number which is the most recent)
    tasks.forEach(task => {
      if (uniqueTasks.has(task.name)) {
        // Compare rows to keep the most recent (highest row number)
        const existingRow = uniqueTasks.get(task.name);
        if (task.row > existingRow) {
          // This is more recent, mark the previous one as duplicate
          duplicates.push(existingRow);
          uniqueTasks.set(task.name, task.row);
        } else {
          // This is older, mark it as duplicate
          duplicates.push(task.row);
        }
      } else {
        uniqueTasks.set(task.name, task.row);
      }
    });
    
    // Sort duplicates in descending order to avoid row shifting issues when deleting
    duplicates.sort((a, b) => b - a);
    
    // Delete duplicates
    const sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
    duplicates.forEach(row => {
      sheet.deleteRow(row);
    });
    
    ui.alert(
      'Cleanup Complete',
      `Removed ${duplicates.length} duplicate tasks.`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    console.error('Error cleaning up duplicates:', error);
    ui.alert('Error', 'Failed to clean up duplicates: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Process a single email message and create a task from it
 * @param {GmailMessage} message - The Gmail message to process
 * @returns {Object} Result of the processing
 */
function processEmailMessage(message) {
  try {
    // Check if we already parsed task details
    if (message.taskDetails) {
      // Use the pre-parsed task details
      const taskDetails = message.taskDetails;
      console.log(`Using pre-parsed task details: ${JSON.stringify(taskDetails)}`);
      
      // Extract individual properties
      const taskName = taskDetails.name;
      let priority = taskDetails.priority;
      const timeBlock = taskDetails.timeBlock;
      const notes = taskDetails.notes;
      
      // DIRECT CHECK: Check if the original subject had a follow-up marker
      const originalSubject = message.getSubject() || '';
      const hasFollowUpMarker = originalSubject.includes('#fup') || 
                               originalSubject.includes('#FUP') || 
                               originalSubject.includes('#Fup');
      
      console.log(`Original subject: "${originalSubject}", Has follow-up marker: ${hasFollowUpMarker}`);
      
      // Normalize follow-up priority to ensure consistent format (lowercase 'u')
      if (String(priority || '').toLowerCase().includes('follow') || hasFollowUpMarker) {
        priority = 'Follow-up'; // Ensure lowercase 'u'
        console.log(`Normalized priority to "Follow-up"`);
      }
      
      // Check if this is a follow-up task (case-insensitive)
      const isFollowUp = String(priority || '').toLowerCase() === 'follow-up' || 
                         String(priority || '').toLowerCase() === 'followup' ||
                         String(priority || '').toLowerCase() === 'follow up' ||
                         hasFollowUpMarker;
      
      console.log(`Task type: ${isFollowUp ? 'Follow-up' : 'Regular'}, Priority: "${priority}", HasMarker: ${hasFollowUpMarker}`);
      
      // Add task to sheet with individual parameters
      const result = sheetManager.addTask(taskName, priority, timeBlock, notes);
      
      // Store the task name in a temporary property for calendar scheduling
      PropertiesService.getScriptProperties().setProperty('lastTaskName', taskName);
      console.log(`Stored task name "${taskName}" for calendar scheduling`);
      
      // For follow-up tasks, double-check the priority was set correctly
      if (isFollowUp) {
        try {
          // Get the sheet and find the priority column
          const sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
          const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          const priorityColIndex = headers.findIndex(header => 
            String(header).toLowerCase() === 'priority'
          );
          
          if (priorityColIndex >= 0) {
            // Force update the priority directly to ensure it's set
            // Use the exact format from the data validation dropdown
            sheet.getRange(result.row, priorityColIndex + 1).setValue('Follow-up');
            console.log(`Ensured Follow-up priority for task at row ${result.row}`);
          }
        } catch (priorityError) {
          console.error('Error setting follow-up priority:', priorityError);
        }
      }
      
      // Send appropriate confirmation email based on task type
      const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
      if (userEmail) {
        // FORCE CORRECT EMAIL TYPE based on multiple checks
        const shouldSendFollowUpEmail = isFollowUp || hasFollowUpMarker || 
                                       String(priority || '').toLowerCase().includes('follow');
        
        console.log(`Email decision: shouldSendFollowUpEmail=${shouldSendFollowUpEmail}, based on isFollowUp=${isFollowUp}, hasFollowUpMarker=${hasFollowUpMarker}`);
        
        if (shouldSendFollowUpEmail) {
          // Send follow-up confirmation with updated subject line
          console.log(`Sending follow-up email notification for "${taskName}"`);
          MailApp.sendEmail({
            to: userEmail,
            subject: `Follow-up Created: ${taskName}`,
            body: `A new follow-up task has been created:

Task: ${taskName}
Priority: ${priority}
Time Block: ${timeBlock} minutes

From: ${message.getFrom()}
Email Link: ${message.getThread().getPermalink()}

This is an automated message from your Digital Assistant.`
          });
        } else {
          // If auto-schedule is enabled, we'll let the scheduler send the email
          // Otherwise, send a task creation confirmation
          const autoSchedule = PropertiesService.getScriptProperties().getProperty('emailTaskAutoSchedule') === 'true';
          
          if (!autoSchedule) {
            // Send regular task confirmation with updated subject line
            console.log(`Sending regular task email notification for "${taskName}"`);
            MailApp.sendEmail({
              to: userEmail,
              subject: `Task Created: ${taskName}`,
              body: `A new task has been created:

Task: ${taskName}
Priority: ${priority}
Time Block: ${timeBlock} minutes

From: ${message.getFrom()}
Email Link: ${message.getThread().getPermalink()}

This is an automated message from your Digital Assistant.`
            });
          }
        }
      }
      
      // If auto-schedule is enabled, schedule the task
      const autoSchedule = PropertiesService.getScriptProperties().getProperty('emailTaskAutoSchedule') === 'true';
      if (autoSchedule && !isFollowUp) {
        try {
          // Initialize task manager if needed
          if (!taskManager.settings) {
            taskManager.initialize();
          }
          
          // Get the newly added task
          const task = sheetManager.getTaskByRow(result.row);
          
          // Schedule the task, passing the task name as a fallback
          taskManager.processTask(task, taskName);
        } catch (scheduleError) {
          console.error('Error scheduling task from email:', scheduleError);
          // Continue processing - we've already created the task
        }
      }
      
      // Mark the message with a label
      const processedLabel = getOrCreateLabel('TasksProcessed');
      message.getThread().addLabel(processedLabel);
      
      return {
        success: true,
        taskName: taskName,
        priority: priority,
        timeBlock: timeBlock,
        row: result.row
      };
    } else {
      // Original code path for direct processing
      // Get message details
      const subject = message.getSubject() || '';
      const sender = message.getFrom() || '';
      const threadId = message.getThread().getId();
      const emailLink = `https://mail.google.com/mail/u/0/#all/${threadId}`;
      
      console.log(`Processing email: "${subject}" from ${sender}`);
      
      // DIRECT CHECK for follow-up markers in the subject
      const hasFollowUpMarker = subject.includes('#fup') || 
                               subject.includes('#FUP') || 
                               subject.includes('#Fup');
      
      console.log(`Original subject: "${subject}", Has follow-up marker: ${hasFollowUpMarker}`);
      
      // Extract task details from subject
      let taskName = subject;
      let priority = 'P2'; // Default priority
      let timeBlock = 30;  // Default time block
      let isFollowUp = hasFollowUpMarker; // Set based on direct check
      
      // Clean up the task name by removing all tags
      // First, save the original subject for tag extraction
      const originalSubject = subject;
      
      // Check for priority markers in subject
      if (originalSubject.includes('#P1') || originalSubject.includes('#p1')) {
        priority = 'P1';
        taskName = taskName.replace(/#P1|#p1/g, '').trim();
      } else if (originalSubject.includes('#P2') || originalSubject.includes('#p2')) {
        priority = 'P2';
        taskName = taskName.replace(/#P2|#p2/g, '').trim();
      } else if (originalSubject.includes('#P3') || originalSubject.includes('#p3')) {
        priority = 'P3';
        taskName = taskName.replace(/#P3|#p3/g, '').trim();
      } else if (originalSubject.includes('#FUP') || originalSubject.includes('#fup') || originalSubject.includes('#Fup')) {
        priority = 'Follow-up'; // Lowercase 'u' to match spreadsheet validation
        isFollowUp = true; // Mark as follow-up task
        taskName = taskName.replace(/#FUP|#fup|#Fup/g, '').trim();
        console.log(`Detected follow-up task: "${taskName}" with priority "${priority}"`);
      }
      
      // Check for time block markers
      const timeBlockMatch = originalSubject.match(/#(\d+)min/);
      if (timeBlockMatch && timeBlockMatch[1]) {
        timeBlock = parseInt(timeBlockMatch[1], 10);
        taskName = taskName.replace(/#\d+min/g, '').trim();
      }
      
      // Remove any other tags
      taskName = taskName.replace(/#\w+/g, '').trim();
      
      // Clean up any double spaces
      taskName = taskName.replace(/\s+/g, ' ').trim();
      
      // Create notes with email link
      const notes = `From: ${sender}\nEmail Link: ${emailLink}`;
      
      // Log the task details before adding to sheet
      console.log(`Adding task to sheet with: Name="${taskName}", Priority="${priority}", TimeBlock=${timeBlock}, IsFollowUp=${isFollowUp}`);
      
      // Add task to sheet
      const result = sheetManager.addTask(taskName, priority, timeBlock, notes);
      
      // Store the task name in a temporary property for calendar scheduling
      PropertiesService.getScriptProperties().setProperty('lastTaskName', taskName);
      console.log(`Stored task name "${taskName}" for calendar scheduling`);
      
      // Send appropriate confirmation email based on task type
      const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
      if (userEmail) {
        // FORCE CORRECT EMAIL TYPE based on direct subject check
        console.log(`Email decision: isFollowUp=${isFollowUp}, hasFollowUpMarker=${hasFollowUpMarker}`);
        
        if (isFollowUp || hasFollowUpMarker) {
          // Send follow-up confirmation
          console.log(`Sending follow-up email notification for "${taskName}"`);
          MailApp.sendEmail({
            to: userEmail,
            subject: `Follow-up Created: ${taskName}`,
            body: `A new follow-up task has been created:

Task: ${taskName}
Priority: Follow-up
Time Block: ${timeBlock} minutes

From: ${sender}
Email Link: ${emailLink}

This is an automated message from your Digital Assistant.`
          });
        } else {
          // If auto-schedule is enabled, we'll let the scheduler send the email
          // Otherwise, send a task creation confirmation
          const autoSchedule = PropertiesService.getScriptProperties().getProperty('emailTaskAutoSchedule') === 'true';
          
          if (!autoSchedule) {
            // Send regular task confirmation
            console.log(`Sending regular task email notification for "${taskName}"`);
            MailApp.sendEmail({
              to: userEmail,
              subject: `Task Created: ${taskName}`,
              body: `A new task has been created:

Task: ${taskName}
Priority: ${priority}
Time Block: ${timeBlock} minutes

From: ${sender}
Email Link: ${emailLink}

This is an automated message from your Digital Assistant.`
            });
          }
        }
      }
      
      // If auto-schedule is enabled, schedule the task
      const autoSchedule = PropertiesService.getScriptProperties().getProperty('emailTaskAutoSchedule') === 'true';
      if (autoSchedule && !isFollowUp) {
        try {
          // Initialize task manager if needed
          if (!taskManager.settings) {
            taskManager.initialize();
          }
          
          // Get the newly added task
          const task = sheetManager.getTaskByRow(result.row);
          
          // Schedule the task, passing the task name as a fallback
          taskManager.processTask(task, taskName);
        } catch (scheduleError) {
          console.error('Error scheduling task from email:', scheduleError);
          // Continue processing - we've already created the task
        }
      }
      
      // Mark the message with a label
      const processedLabel = getOrCreateLabel('TasksProcessed');
      message.getThread().addLabel(processedLabel);
      
      return {
        success: true,
        taskName: taskName,
        priority: priority,
        timeBlock: timeBlock,
        row: result.row
      };
    }
  } catch (error) {
    console.error('Error processing email message:', error);
    
    // Mark the message with an error label
    try {
      const errorLabel = getOrCreateLabel('TasksError');
      message.getThread().addLabel(errorLabel);
    } catch (labelError) {
      console.error('Error adding error label:', labelError);
    }
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Debug function to test priority extraction from email subjects
 */
function testEmailPriorityExtraction() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Test Email Priority Extraction',
    'Enter an email subject to test:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const subject = response.getResponseText().trim();
    
    let taskName = subject;
    let priority = 'P2'; // Default priority
    const originalSubject = subject;
    
    // Check for priority markers in subject
    if (originalSubject.includes('#P1') || originalSubject.includes('#p1')) {
      priority = 'P1';
      taskName = taskName.replace(/#P1|#p1/g, '').trim();
    } else if (originalSubject.includes('#P2') || originalSubject.includes('#p2')) {
      priority = 'P2';
      taskName = taskName.replace(/#P2|#p2/g, '').trim();
    } else if (originalSubject.includes('#P3') || originalSubject.includes('#p3')) {
      priority = 'P3';
      taskName = taskName.replace(/#P3|#p3/g, '').trim();
    } else if (originalSubject.includes('#FUP') || originalSubject.includes('#fup') || originalSubject.includes('#Fup')) {
      priority = 'Follow-up'; // Lowercase 'u' to match spreadsheet validation
      taskName = taskName.replace(/#FUP|#fup|#Fup/g, '').trim();
    }
    
    // Check for time block markers
    const timeBlockMatch = originalSubject.match(/#(\d+)min/);
    if (timeBlockMatch && timeBlockMatch[1]) {
      taskName = taskName.replace(/#\d+min/g, '').trim();
    }
    
    // Remove any other tags that might be present (starting with #)
    taskName = taskName.replace(/#\w+/g, '').trim();
    
    // Clean up any double spaces
    taskName = taskName.replace(/\s+/g, ' ').trim();
    
    // Remove any JSON-like formatting that might be in the subject
    taskName = taskName.replace(/\{.*?\}/g, '').trim();
    
    ui.alert('Result', 
      `Original: "${subject}"
       Cleaned Task Name: "${taskName}"
       Priority: "${priority}"`, 
      ui.ButtonSet.OK);
  }
}

/**
 * Find emails with task markers
 * @returns {Array} Array of Gmail messages
 */
function findTaskEmails() {
  try {
    // ... existing code ...
    
    // Process each thread
    const taskMessages = [];
    threads.forEach(thread => {
      // ... existing code ...
      
      // Check if any message in the thread has a task marker
      let hasTaskMarker = false;
      let taskMessage = null;
      
      messages.forEach(message => {
        const subject = message.getSubject() || '';
        
        // Check for task markers in subject
        if (subject.includes('#td') || 
            subject.includes('#TD') || 
            subject.includes('#fup') || 
            subject.includes('#FUP') || 
            subject.includes('#Fup')) {
          hasTaskMarker = true;
          taskMessage = message;
          
          // Parse task details from subject
          const taskDetails = parseTaskFromSubject(subject, message);
          console.log('Parsed task details:', taskDetails);
          
          // Explicitly check if this is a follow-up task
          const isFollowUp = subject.includes('#fup') || 
                            subject.includes('#FUP') || 
                            subject.includes('#Fup');
          
          if (isFollowUp) {
            console.log(`Explicitly marking as follow-up task: ${taskDetails.name}`);
            // Ensure priority is set to Follow-up
            taskDetails.priority = 'Follow-up';
          }
          
          // Store task details in message properties for later use
          message.taskDetails = taskDetails;
        }
      });
      
      // ... rest of the function ...
    });
    
    return taskMessages;
  } catch (error) {
    // ... error handling ...
  }
}

/**
 * Parse task details from email subject
 * @param {string} subject - Email subject
 * @param {GmailMessage} message - Gmail message
 * @returns {Object} Task details
 */
function parseTaskFromSubject(subject, message) {
  let taskName = subject;
  let priority = 'P2'; // Default priority
  let timeBlock = 30;  // Default time block
  
  // Clean up the task name by removing all tags
  const originalSubject = subject;
  
  // Check for priority markers
  if (originalSubject.includes('#P1') || originalSubject.includes('#p1')) {
    priority = 'P1';
    taskName = taskName.replace(/#P1|#p1/g, '').trim();
  } else if (originalSubject.includes('#P2') || originalSubject.includes('#p2')) {
    priority = 'P2';
    taskName = taskName.replace(/#P2|#p2/g, '').trim();
  } else if (originalSubject.includes('#P3') || originalSubject.includes('#p3')) {
    priority = 'P3';
    taskName = taskName.replace(/#P3|#p3/g, '').trim();
  } else if (originalSubject.includes('#FUP') || originalSubject.includes('#fup') || originalSubject.includes('#Fup')) {
    priority = 'Follow-up'; // Lowercase 'u' to match spreadsheet validation
    taskName = taskName.replace(/#FUP|#fup|#Fup/g, '').trim();
    console.log(`Parsed FUP task name: ${taskName}`);
  }
  
  // Check for time block markers
  const timeBlockMatch = originalSubject.match(/#(\d+)min/);
  if (timeBlockMatch && timeBlockMatch[1]) {
    timeBlock = parseInt(timeBlockMatch[1], 10);
    taskName = taskName.replace(/#\d+min/g, '').trim();
  }
  
  // Remove any other tags
  taskName = taskName.replace(/#\w+/g, '').trim();
  
  // Clean up any double spaces
  taskName = taskName.replace(/\s+/g, ' ').trim();
  
  // Create notes with email link
  const threadId = message.getThread().getId();
  const emailLink = `https://mail.google.com/mail/u/0/#all/${threadId}`;
  const notes = `Email Link: ${emailLink}`;
  
  // Create task object
  const task = {
    name: taskName,
    priority: priority,
    timeBlock: timeBlock,
    notes: notes,
    status: 'Pending'
  };
  
  console.log(`Successfully parsed task: ${JSON.stringify(task)}`);
  return task;
}

/**
 * Test function to send a follow-up task notification email
 */
function testFollowUpEmailNotification() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
    if (!userEmail) {
      ui.alert('Error', 'User email not configured. Please set it in script properties.', ui.ButtonSet.OK);
      return;
    }
    
    // Send a test follow-up notification
    MailApp.sendEmail({
      to: userEmail,
      subject: `Follow-up Created: Test Follow-up Task`,
      body: `A new follow-up task has been created:

Task: Test Follow-up Task
Priority: Follow-up
Time Block: 30 minutes

From: Test Sender
Email Link: https://mail.google.com/mail/u/0/#all/test

This is an automated message from your Digital Assistant.`
    });
    
    ui.alert('Success', 'Test follow-up notification sent. Please check your email.', ui.ButtonSet.OK);
  } catch (error) {
    console.error('Error sending test notification:', error);
    ui.alert('Error', 'Failed to send test notification: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Debug function to test email notification for follow-up tasks
 */
function debugEmailNotification() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Create a mock message object
    const mockMessage = {
      getSubject: () => "#fup Debug Test Task",
      getFrom: () => "test@example.com",
      getThread: () => ({
        getId: () => "debug123",
        getPermalink: () => "https://mail.google.com/mail/u/0/#all/debug123",
        addLabel: () => console.log("Label added")
      })
    };
    
    // Process the mock message
    console.log("Processing mock follow-up message");
    const result = processEmailMessage(mockMessage);
    
    // Show result
    ui.alert("Debug Result", 
             `Task processed: ${result.success ? "Success" : "Failed"}\n` +
             `Task name: ${result.taskName}\n` +
             `Priority: ${result.priority}\n` +
             `Row: ${result.row}`, 
             ui.ButtonSet.OK);
    
  } catch (error) {
    console.error("Debug error:", error);
    ui.alert("Error", "Debug failed: " + error.message, ui.ButtonSet.OK);
  }
} 