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

  // --- Track processed thread IDs ---
  const processedThreadIds = new Set();
  // ---------------------------------

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
      // ---> Limit to 1 hour <---
      // const taskEmailSearch = `from:${userEmail} newer_than:15m -label:Task-Processed -label:Task-Error-Processed`;
      
      // console.log('Using search query:', taskEmailSearch);

      // --- Define search queries ---
      // Query for standard #td/#fup tasks (last 1 hour)
      const standardTaskSearch = `from:${userEmail} newer_than:1h -label:Task-Processed -label:Task-Error-Processed -label:DailyTaskReply-Processed (subject:"#td" OR subject:"#fup")`;
      
      // Query for replies to the daily prompt (last 1 hour)
      const dailyReplySearchSubject = "Re: What tasks will you tackle today? [Daily Task Entry]";
      // ---> Limit to 1 hour <---
      const dailyReplySearch = `from:${userEmail} newer_than:1h subject:("${dailyReplySearchSubject}") -label:DailyTaskReply-Processed -label:Task-Processed -label:Task-Error-Processed`;
      
      // ---> Corrected Logging <---
      console.log('Standard Task Query:', standardTaskSearch);
      console.log('Daily Reply Query:', dailyReplySearch);
      
      // Search for matching emails - THIS GETS ALL RECENT UNPROCESSED EMAILS
      // const allRecentThreads = GmailApp.search(taskEmailSearch);
      // console.log(`Found ${allRecentThreads.length} potentially relevant email threads`);

      // --- Search for standard tasks only ---
      const standardThreads = GmailApp.search(standardTaskSearch);
      // const dailyReplyThreads = GmailApp.search(dailyReplySearch); // DISABLED daily reply search

      console.log(`Found ${standardThreads.length} potential standard task emails.`);
      // console.log(`Found ${dailyReplyThreads.length} potential daily task replies.`); // DISABLED logging
      
      // Combine threads, ensuring no duplicates - NOW ONLY USES standardThreads
      const allThreads = [...standardThreads];
      // const standardThreadIds = new Set(standardThreads.map(t => t.getId())); // DISABLED logic for combining
      // dailyReplyThreads.forEach(thread => { // DISABLED logic for combining
      //    if (!standardThreadIds.has(thread.getId())) { // DISABLED logic for combining
      //       allThreads.push(thread); // DISABLED logic for combining
      //    }
      // }); // DISABLED logic for combining
      const totalThreadsFound = allThreads.length;
      console.log(`Processing a total of ${totalThreadsFound} unique standard threads found.`);
      
      // --- REMOVE Filtering by subject here, process all threads ---
      // const taskThreads = allRecentThreads.filter(thread => { ... });
      // console.log(`Found ${taskThreads.length} threads with task markers (#td or #fup)`);
      // -------------------------------------------------------------
      
      // Track processed tasks names within this run (if needed, less critical now)
      const processedTaskNames = new Set();
      
      // ---> ADD Processing Limit <---
      const MAX_THREADS_PER_RUN = 50; // Limit to prevent quota errors
      let threadsProcessedThisRun = 0;
      // ---> END Processing Limit <---

      // Process each thread found by the search
      for (const thread of allThreads) {
          // ---> Check Processing Limit <---
          if (threadsProcessedThisRun >= MAX_THREADS_PER_RUN) {
            console.log(`Reached processing limit of ${MAX_THREADS_PER_RUN} threads for this run. Halting processing.`);
            break; // Exit the loop
          }
          // ---> Increment Counter <---
          threadsProcessedThisRun++;
          // ---> END Check <---

          const threadId = thread.getId();

          // --- Skip if already processed in this run ---
          if (processedThreadIds.has(threadId)) {
            continue; 
          }
          // -------------------------------------------

          const messages = thread.getMessages();
          
          // ----> ADD CHECK FOR EMPTY MESSAGES <----
          if (!messages || messages.length === 0) {
            console.log(`Skipping thread ${threadId} because it contains no messages.`);
            // Mark as processed to avoid re-checking this thread unnecessarily
            processedThreadIds.add(threadId);
            // Optionally, apply a specific label like 'Empty-Thread' if needed for tracking
            continue;
          }
          // ----> END CHECK <----

          const message = messages[messages.length - 1];

          // ----> ADD CHECK FOR VALID MESSAGE OBJECT <----
          if (!message) {
              console.log(`Skipping thread ${threadId} as the last message object is invalid or couldn't be retrieved.`);
              processedThreadIds.add(threadId); // Mark as processed
              continue;
          }
          // ----> END CHECK <----

          console.log(`Processing Thread ID: ${threadId}`);

          try {
            // --- Get subject and body ONCE INSIDE the try block ---
            const subject = message.getSubject() || '';
            const plainBody = message.getPlainBody() || ''; // Fetch body once
            const emailUrl = `https://mail.google.com/mail/u/0/#all/${threadId}`;
            console.log(`Subject: "${subject}"`);

            // --- CHECK 1: Is this a reply to the daily prompt? --- // UNCOMMENTING BLOCK
            
            if (subject.includes(dailyReplySearchSubject)) {
               // ----> SIMPLIFIED DEBUG LOGGING <----
               console.log(`DEBUG: Preparing to call parseAndAddDailyTasksFromReply for thread ${threadId}`); // Log is a bit misleading now as it's inlined
               try {
                 // Log only basic info, avoid method calls on 'message' initially
                 console.log(`DEBUG: typeof message: ${typeof message}`);
                 console.log(`DEBUG: message object (raw):`, message);
                 // Check if it has the method *without* calling it
                 if (message && typeof message.getPlainBody === 'function') {
                    console.log(`DEBUG: message appears to have getPlainBody method.`);
                 } else {
                    console.warn(`DEBUG: message does NOT appear to have getPlainBody method.`);
                 }
               } catch(debugError) {
                 // Explicitly log if the DEBUG block itself fails
                 console.error(`>>> DEBUG BLOCK FAILED: Error accessing message properties before call: ${debugError}`, debugError.stack);
               }
               // ----> END SIMPLIFIED DEBUG LOGGING <----

               // ----> Final check (keep this, but log clearly) <----
               if (!message || typeof message.getPlainBody !== 'function') {
                   console.error(`CRITICAL CHECK FAILED (Before Inline Parse): Message object invalid for thread ${threadId}. Type: ${typeof message}, Has getPlainBody func: ${typeof message?.getPlainBody === 'function'}. Skipping.`);
                   // Mark as processed with error label to be safe
                   ensureLabelExists('Task-Error-Processed');
                   thread.addLabel(GmailApp.getUserLabelByName('Task-Error-Processed'));
                   processedThreadIds.add(threadId);
                   continue;
               }
               // ----> END FINAL CHECK <----

               console.log(`Detected Daily Task Reply: ${threadId}. Processing inline...`);
               const messageToParse = message; // Use the message we already validated
               let tasksAddedInline = 0;

               // ----> START INLINE LOGIC (from parseAndAddDailyTasksFromReply) <----
               // Use the 'plainBody' variable fetched ONCE at the start of the outer try block
               let inlinePermalink = '';
               let inlineMessageId = '';

               try {
                   // Directly use messageToParse here for ID/Permalink
                   inlinePermalink = messageToParse.getThread().getPermalink();
                   inlineMessageId = messageToParse.getId();
                   console.log(`INLINE: Processing message ID: ${inlineMessageId} in thread: ${inlinePermalink} using pre-fetched body.`);

                   // ---> USE PRE-FETCHED BODY (plainBody) <--- 
                   // const inlineBody = plainBody; // This is already available as 'plainBody'

                   // Now, process the plainBody (which might be empty if retrieval failed or returned null/undefined)
                   const inlineLines = plainBody.split('\\n');
                   const taskRegex = /^\\s*\\d+\\s*-\\s*(.+?)\\s*,\\s*(\\d+)\\s*(?:Mins|Min|Minutes|Minute)/i;

                   for (const line of inlineLines) {
                       const trimmedLine = line.trim();
                       if (!trimmedLine) continue;
                       const match = trimmedLine.match(taskRegex);
                       if (match && match[1] && match[2]) {
                           const taskName = match[1].trim();
                           const timeBlock = parseInt(match[2], 10);
                           if (taskName && !isNaN(timeBlock) && timeBlock > 0) {
                               console.log(`INLINE Parsed daily task from message ${inlineMessageId}: Name="${taskName}", Time=${timeBlock}`);
                               try {
                                   const task = {
                                       name: taskName, priority: 'P1', timeBlock: timeBlock,
                                       notes: `Added from daily email reply. Original email: ${inlinePermalink}`,
                                       status: 'Pending'
                                   };
                                   const addedTask = sheetManager.addTask(task);
                                   if (addedTask && addedTask.success) {
                                       console.log(`INLINE Added daily task "${taskName}" to sheet.`);
                                       tasksAddedInline++;
                                   } else {
                                       console.warn(`INLINE Failed to add daily task "${taskName}" (from message ${inlineMessageId}) to sheet.`);
                                   }
                               } catch (e) {
                                   console.error(`INLINE Error adding daily task "${taskName}" (from message ${inlineMessageId}) to sheet: ${e}`);
                               }
                           } else {
                               console.log(`INLINE Skipping line in message ${inlineMessageId} (invalid format or data): "${trimmedLine}"`);
                           }
                       } else {
                           console.log(`INLINE Skipping line in message ${inlineMessageId} (did not match regex): "${trimmedLine}"`);
                       }
                   }

                   // Check if tasks were added and handle labeling/archiving
                   if (tasksAddedInline > 0) {
                       console.log(`INLINE: Successfully added ${tasksAddedInline} tasks.`);
                       // Send confirmation email
                       try {
                           MailApp.sendEmail({
                               to: userEmail,
                               subject: `Daily Tasks Added: ${tasksAddedInline} tasks created (Inline Process)`,
                               body: `Successfully added ${tasksAddedInline} tasks from your daily reply.\\n\\nEmail Link: ${inlinePermalink}`
                           });
                       } catch (emailError) {
                           console.error(`INLINE Error sending daily task confirmation email: ${emailError}`);
                       }
                       // Mark as processed (use the specific daily label)
                       messageToParse.markRead();
                       ensureLabelExists('DailyTaskReply-Processed');
                       thread.addLabel(GmailApp.getUserLabelByName('DailyTaskReply-Processed'));
                       processedThreadIds.add(threadId);
                       // Archive
                       if (thread.isInInbox()) {
                           thread.moveToArchive();
                       }
                   } else {
                       console.log(`INLINE: No valid tasks found in the body of message ${inlineMessageId}. Marking processed.`);
                       // Mark as processed even if no tasks found to avoid retries
                       messageToParse.markRead();
                       ensureLabelExists('DailyTaskReply-Processed');
                       thread.addLabel(GmailApp.getUserLabelByName('DailyTaskReply-Processed'));
                       processedThreadIds.add(threadId);
                   }
               } catch (inlineProcessingError) {
                   console.error(`INLINE ERROR processing daily reply. Message ID ${inlineMessageId || 'unknown'}. Error: ${inlineProcessingError}`, inlineProcessingError.stack);
                   // Apply error label if inline processing fails
                   ensureLabelExists('Task-Error-Processed');
                   thread.addLabel(GmailApp.getUserLabelByName('Task-Error-Processed'));
                   processedThreadIds.add(threadId);
               }
               // ----> END INLINE LOGIC <----

               continue; // Skip standard processing for this thread since it was handled inline
            }
            
            // --- END UNCOMMENTED DAILY REPLY BLOCK ---
            
            // --- Process Standard Task (#td / #fup) --- (This block is now always executed if not skipped earlier)
            console.log(`Processing as Standard Task (#td/#fup): ${threadId}`);
            try {
              let task; // Declare task variable outside the if/else
              const lowerCaseSubject = subject.toLowerCase();

              // ---> ADDED: Specific log before decision <---
              console.log(`DEBUG: Checking subject "${lowerCaseSubject}". Contains #fup? ${lowerCaseSubject.includes('#fup')}. Contains #td? ${lowerCaseSubject.includes('#td')}.`);

              // ---> ADDED: Check if #fup or #td <---
              if (lowerCaseSubject.includes('#fup')) {
                console.log('Detected #fup, calling parseFollowUpFormat');
                task = parseFollowUpFormat(subject, emailUrl);
              } else if (lowerCaseSubject.includes('#td')) { // Also check for #td explicitly
                console.log('Detected #td, calling parseTaskFormat');
                task = parseTaskFormat(subject, emailUrl, plainBody);
              } else {
                // This case should technically not be hit due to the initial search query,
                // but include it defensively.
                console.warn(`Subject "${subject}" matched general search but lacks specific #fup or #td marker. Skipping parsing.`);
                task = null; // Ensure task is null if neither marker is found
              }
              // ---> END ADDED CHECK <---
              
              if (task) {
                // --- Add Email Origin Info ---
                task.origin = 'email';
                task.emailThreadId = threadId;
                task.emailMessageId = message.getId();
                task.emailSubject = subject; // Store original subject if needed later
                // --- End Add Email Origin Info ---
                
                console.log('Task object received:', JSON.stringify(task)); // More generic log here
                
                // Use the task object returned by the correct parser
                const addedRow = sheetManager.addTask(task); 
                
                if (addedRow) {
                  console.log(`Added task "${task.name}" to sheet row ${addedRow}. Status: ${task.status}.`);
                  processedTaskNames.add(task.name); // Track processed name
                  // Apply processed label immediately after successful sheet add
                  ensureLabelExists('Task-Processed');
                  thread.addLabel(GmailApp.getUserLabelByName('Task-Processed'));
                } else {
                  console.error(`Failed to add task "${task.name}" to sheet.`);
                  // Apply error label if sheet add fails
                  ensureLabelExists('Task-Error-Processed');
                  thread.addLabel(GmailApp.getUserLabelByName('Task-Error-Processed'));
                }
              } else {
                console.warn(`Could not parse task from subject: "${subject}". Skipping thread ${threadId}.`);
                // Apply error label if parsing fails
                ensureLabelExists('Task-Error-Processed');
                thread.addLabel(GmailApp.getUserLabelByName('Task-Error-Processed'));
              }
            } catch (taskParseError) {
                console.error(`Error parsing/adding standard task from thread ${threadId}: ${taskParseError}`, taskParseError.stack);
                ensureLabelExists('Task-Error-Processed');
                thread.addLabel(GmailApp.getUserLabelByName('Task-Error-Processed'));
            }
            // Mark thread as processed (success or failure) for standard tasks
            processedThreadIds.add(threadId); 

          } catch (messageProcessingError) {
            console.error(`Error processing thread ID ${threadId}:`, messageProcessingError.message, messageProcessingError.stack);
            
            // Send error notification 
            sendErrorNotification(subject, messageProcessingError.message, userEmail);
            
            // Mark as error-processed
            try {
               message.markRead();
               ensureLabelExists('Task-Error-Processed');
               thread.addLabel(GmailApp.getUserLabelByName('Task-Error-Processed'));
               processedThreadIds.add(threadId); // Also mark errors as processed
            } catch (labelError) {
               console.error(`Failed to apply error label to thread ${threadId}: ${labelError}`);
            }
          }
        } // End for loop processing threads
    } catch (error) {
      // Catch errors in the main processing block (e.g., getting sheet data)
      console.error('Error during email processing setup or loop:', error);
    }
    
    // Check spreadsheet state at end
    try {
      const endRowCount = SpreadsheetApp.getActive().getSheetByName('Tasks').getLastRow();
      console.log(`PROCESS END: Spreadsheet has ${endRowCount} rows at end (change: ${endRowCount - startRowCount})`);
    } catch (e) {
      console.error("Error getting end row count: " + e);
    }
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
  const prefix = '#fup';
  const prefixLength = prefix.length;
  let remainingPart = ''; // Initialize remainingPart

  if (subject.toLowerCase().startsWith(prefix)) {
    remainingPart = subject.substring(prefixLength);
  } else {
    // Find the index if #fup is not at the start
    const index = subject.toLowerCase().indexOf(prefix);
    if (index !== -1) {
      remainingPart = subject.substring(index + prefixLength);
    } else {
      // Fallback if #fup isn't found (shouldn't happen if called correctly, but safe)
      remainingPart = subject; 
    }
  }

  // Remove leading colon, space, or hyphen if present after the prefix
  if (remainingPart.startsWith(':') || remainingPart.startsWith(' ') || remainingPart.startsWith('-')) {
     taskName = remainingPart.substring(1).trim();
  } else {
     taskName = remainingPart.trim();
  }
  
  // Clean up task name by removing any Fwd: or Re: prefixes
  taskName = taskName.replace(/^(Fwd|Re|FWD|RE):\\s*/i, '').trim();
  
  console.log('Parsed FUP task name:', taskName);
  
  return {
    name: taskName,
    priority: 'Follow-up', // Priority should be Follow-up
    timeBlock: 30, // Default time block for follow-ups
    notes: `Email Link: ${emailUrl}`,
    status: 'Follow-up' // Status should be Follow-up
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
  // Check if the label exists
  if (!GmailApp.getUserLabelByName(labelName)) {
    try {
      // Create the label if it doesn't exist
      GmailApp.createLabel(labelName);
      console.log(`Created Gmail label: ${labelName}`);
    } catch (error) {
      // Handle potential race condition or other error during creation
      console.error(`Error creating label ${labelName}: ${error.message}`);
      // Check again in case it was created by another process
      if (!GmailApp.getUserLabelByName(labelName)) {
         console.error(`Failed to ensure label ${labelName} exists.`);
      }
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

/**
 * Sends a daily email prompt asking for tasks for the day.
 * Should be triggered by a time-driven trigger daily around 10 AM.
 */
function sendDailyTaskPrompt() {
  const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
  if (!userEmail) {
    console.error('User email not configured for daily prompt. Please run setup first.');
    return;
  }

  const subject = "What tasks will you tackle today? [Daily Task Entry]";
  const body = `Reply to this email with your tasks for today, one per line, like this:

1 - Task Name 1, 30 Mins
2 - Task Name 2, 60 Mins

The script will add these as P1 tasks to your sheet.
`;

  try {
    MailApp.sendEmail({
      to: userEmail,
      subject: subject,
      body: body
    });
    console.log(`Daily task prompt email sent to ${userEmail}`);
  } catch (error) {
    console.error(`Error sending daily task prompt email: ${error.message}`);
  }
}

/**
 * Parses tasks from the body of a reply to the daily prompt email.
 * @param {GmailMessage} emailMessageObject - The reply email message.
 * @param {string} userEmail - The user's email address.
 * @returns {boolean} True if tasks were successfully parsed and added, false otherwise.
 */
function parseAndAddDailyTasksFromReply(emailMessageObject, userEmail) {
  // ----> Log entry and received parameter immediately <----
  console.log(`INFO: Entered parseAndAddDailyTasksFromReply.`);
  // ----> Log the RENAMED parameter <----
  console.log(`INFO: Received emailMessageObject parameter (raw):`, emailMessageObject);
  console.log(`INFO: typeof emailMessageObject parameter: ${typeof emailMessageObject}`);
  // ----> END Log entry <----
  
  // ----> Refined Initial Check using RENAMED parameter <----
  if (!emailMessageObject || typeof emailMessageObject.getPlainBody !== 'function') {
    console.error(`CRITICAL CHECK FAILED (Inside Function): Invalid message object received. Type: ${typeof emailMessageObject}, Has getPlainBody func: ${typeof emailMessageObject?.getPlainBody === 'function'}. Cannot process.`);
    // Optionally log thread ID if possible
    try { console.error(`Parent Thread ID (if available): ${emailMessageObject.getThread().getId()}`); } catch (e) {}
    return false; // Cannot proceed
  }
  // ----> END Refined Check <----

  console.log('Attempting to parse tasks from daily prompt reply.');
  let body = ''; // Initialize body
  let threadPermalink = ''; // Initialize permalink
  let messageId = ''; // Initialize message ID for logging

  try {
    // Use RENAMED parameter
    threadPermalink = emailMessageObject.getThread().getPermalink();
    messageId = emailMessageObject.getId();

    console.log(`Processing message ID: ${messageId} in thread: ${threadPermalink}`);

    // ---> TRY getting the body specifically using RENAMED parameter <---
    body = emailMessageObject.getPlainBody();

    // Handle cases where getPlainBody returns null/undefined instead of throwing
    if (body === null || body === undefined) {
        console.warn(`Message body retrieved as null or undefined for message ${messageId}. Treating as empty.`);
        body = ''; // Ensure body is a string for split
    }

  } catch (e) {
    console.error(`ERROR calling getPlainBody() for message ID ${messageId}: ${e.message}`);
    // Attempt to log more details about the specific message object that failed
    try {
      // Use RENAMED parameter
      console.error(`Failed Message Details: Subject='${emailMessageObject.getSubject()}', From='${emailMessageObject.getFrom()}', Date='${emailMessageObject.getDate()}', ID='${messageId}', Permalink='${threadPermalink}'`);
    } catch (logError) {
      console.error(`Could not retrieve additional details for the failed message ${messageId}.`);
    }
    return false; // Cannot proceed without body
  }

  // Now, process the body (which might be empty if retrieval failed or returned null/undefined)
  const lines = body.split('\\n');
  let tasksAdded = 0;
  const taskRegex = /^\\s*\\d+\\s*-\\s*(.+?)\\s*,\\s*(\\d+)\\s*(?:Mins|Min|Minutes|Minute)/i;

  for (const line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine) continue;

    const match = trimmedLine.match(taskRegex);
    if (match && match[1] && match[2]) {
      const taskName = match[1].trim();
      const timeBlock = parseInt(match[2], 10);

      if (taskName && !isNaN(timeBlock) && timeBlock > 0) {
        console.log(`Parsed daily task from message ${messageId}: Name="${taskName}", Time=${timeBlock}`);
        try {
          const task = {
            name: taskName,
            priority: 'P1',
            timeBlock: timeBlock,
            notes: `Added from daily email reply. Original email: ${threadPermalink}`,
            status: 'Pending'
          };
          const addedTask = sheetManager.addTask(task);
          if (addedTask && addedTask.success) {
            console.log(`Added daily task "${taskName}" to sheet.`);
            tasksAdded++;
          } else {
             console.warn(`Failed to add daily task "${taskName}" (from message ${messageId}) to sheet.`);
          }
        } catch (e) {
          console.error(`Error adding daily task "${taskName}" (from message ${messageId}) to sheet: ${e}`);
        }
      } else {
        console.log(`Skipping line in message ${messageId} (invalid format or data): "${trimmedLine}"`);
      }
    } else {
       console.log(`Skipping line in message ${messageId} (did not match regex): "${trimmedLine}"`);
    }
  }

  if (tasksAdded > 0) {
     try {
       MailApp.sendEmail({
         to: userEmail,
         subject: `Daily Tasks Added: ${tasksAdded} tasks created`,
         body: `Successfully added ${tasksAdded} tasks from your daily reply.\\n\\nEmail Link: ${threadPermalink}`
       });
     } catch (emailError) {
       console.error(`Error sending daily task confirmation email: ${emailError}`);
     }
     return true;
  } else {
     console.log(`No valid tasks found in the body of message ${messageId} (or body was empty/unreadable). Thread: ${threadPermalink}`);
     return false;
  }
} 