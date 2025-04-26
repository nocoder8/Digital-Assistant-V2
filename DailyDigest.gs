/**
 * Daily Digest functionality
 * Provides a daily email summary of tasks
 * 
 * This is a standalone module that can be enabled/disabled independently
 * without affecting other system features.
 */

/**
 * Send daily digest email with task summaries
 * This function can be triggered by a time-based trigger
 */
function sendDailyTaskDigest() {
  console.log('Starting daily task digest');
  
  // Check if feature is enabled
  const isEnabled = PropertiesService.getScriptProperties().getProperty('dailyDigestEnabled') === 'true';
  if (!isEnabled) {
    console.log('Daily digest is disabled. Skipping.');
    return false;
  }
  
  try {
    // Create a safe wrapper around task manager initialization
    // This ensures we don't disrupt other processes if there's an issue
    let taskManagerInitialized = false;
    try {
      // Initialize task manager if needed, but don't throw if it fails
      if (!taskManager.settings) {
        taskManager.initialize();
      }
      taskManagerInitialized = true;
    } catch (initError) {
      console.warn('Could not initialize task manager, continuing with limited functionality:', initError);
    }
    
    // Get user email from settings
    const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
    if (!userEmail) {
      console.error('User email not configured. Cannot send digest.');
      return false;
    }
    
    // Get all tasks - use try/catch to prevent failures
    let allTasks = [];
    try {
      allTasks = sheetManager.getTasks();
      console.log(`Retrieved ${allTasks.length} tasks from sheet`);
      
      // Log all priorities to help diagnose the issue
      const priorities = allTasks.map(task => task.priority);
      console.log('Task priorities found:', JSON.stringify(priorities));
    } catch (sheetError) {
      console.error('Error getting tasks from sheet:', sheetError);
      // Send error notification instead of failing completely
      MailApp.sendEmail({
        to: userEmail,
        subject: `Daily Task Digest - Error`,
        body: `There was an error generating your daily task digest: ${sheetError.message}\n\nPlease check your task sheet.`
      });
      return false;
    }
    
    // 1. Follow-up tasks
    const followUpTasks = allTasks.filter(task => {
      const priority = String(task.priority || '').toLowerCase();
      return priority === 'follow-up' || priority === 'followup' || priority === 'follow up';
    });
    
    console.log(`Found ${followUpTasks.length} follow-up tasks`);
    
    // 2. Upcoming scheduled tasks (tasks with status "Scheduled")
    const scheduledTasks = allTasks.filter(task => task.status === 'Scheduled');
    console.log(`Found ${scheduledTasks.length} scheduled tasks`);
    
    // 3. Error tasks (tasks that failed to schedule)
    const errorTasks = allTasks.filter(task => 
      task.status === 'Error' || 
      task.status === 'No Calendar Access' ||
      task.status === 'Failed'
    );
    console.log(`Found ${errorTasks.length} error tasks`);
    
    // Format follow-up tasks
    let followUpSection = '';
    if (followUpTasks.length > 0) {
      followUpSection = '📋 Follow-up Tasks:\n\n';
      followUpTasks.forEach(task => {
        followUpSection += `• ${task.name} (${task.timeBlock} min)\n`;
      });
      followUpSection += '\n';
    }
    
    // Generate HTML email content
    const htmlContent = generateDigestHtml(followUpTasks, scheduledTasks, errorTasks, taskManagerInitialized);
    
    // Send the email
    MailApp.sendEmail({
      to: userEmail,
      subject: `📊 Daily Task Digest - ${new Date().toLocaleDateString()}`,
      htmlBody: htmlContent
    });
    
    console.log('Daily digest email sent successfully');
    return true;
  } catch (error) {
    console.error('Error sending daily digest:', error);
    
    // Try to send a simple error notification
    try {
      const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
      if (userEmail) {
        MailApp.sendEmail({
          to: userEmail,
          subject: `Daily Task Digest - Error`,
          body: `There was an error generating your daily task digest: ${error.message}`
        });
      }
    } catch (mailError) {
      console.error('Could not send error notification:', mailError);
    }
    
    return false;
  }
}

/**
 * Generate HTML content for the digest email
 */
function generateDigestHtml(followUpTasks, scheduledTasks, errorTasks, hasCalendarAccess) {
  // Warning message if calendar access is not available
  let warningMessage = '';
  if (!hasCalendarAccess) {
    warningMessage = `
      <div style="background-color:#fff3e0; padding:8px; border-left:4px solid #ff9800; margin-bottom:15px; font-size:13px;">
        <strong>⚠️ Limited Information:</strong> Calendar access is not available. 
        Scheduled times may not be accurate.
      </div>
    `;
  }
  
  // Generate follow-up tasks table with improved email link detection
  const followUpHtml = generateFollowUpTasksHtml(followUpTasks);
  
  // Generate scheduled tasks table using the new function
  const scheduledHtml = generateScheduledTasksHtml(scheduledTasks, hasCalendarAccess);
  
  // Generate error tasks table
  const errorHtml = generateErrorTasksHtml(errorTasks);
  
  // Combine all sections (without the summary)
  return `
    <html>
      <head>
        <style>
          .empty-message { color: #7f8c8d; font-style: italic; font-size: 13px; }
        </style>
      </head>
      <body style="font-family:Arial,sans-serif; color:#333; line-height:1.4; max-width:800px; margin:0 auto; padding:20px;">
        <h1 style="color:#2c3e50; border-bottom:1px solid #eee; padding-bottom:10px; margin-bottom:15px; font-size:20px;">📊 Daily Task Digest - ${new Date().toLocaleDateString()}</h1>
        
        ${warningMessage}
        
        <h2 style="color:#3498db; margin-top:15px; margin-bottom:10px; font-size:16px;">📬 Follow-Up Tasks</h2>
        ${followUpHtml}
        
        <h2 style="color:#3498db; margin-top:15px; margin-bottom:10px; font-size:16px;">📅 Upcoming Scheduled Tasks</h2>
        ${scheduledHtml}
        
        <h2 style="color:#3498db; margin-top:15px; margin-bottom:10px; font-size:16px;">⚠️ Tasks with Errors</h2>
        ${errorHtml}
        
        <div style="margin-top:20px; padding-top:10px; border-top:1px solid #eee; font-size:0.9em; color:#7f8c8d;">
          <p>This is an automated message from your Digital Assistant. 🤖</p>
        </div>
      </body>
    </html>
  `;
}

/**
 * Generate follow-up tasks HTML with better email link detection
 */
function generateFollowUpTasksHtml(followUpTasks) {
  if (followUpTasks.length === 0) {
    return '<p class="empty-message">✅ No follow-up tasks.</p>';
  }
  
  return `
    <table cellspacing="0" cellpadding="4" border="1" style="width:100%; border-collapse:collapse; font-size:13px; border:1px solid #ddd;">
      <tr>
        <th width="40%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Task</th>
        <th width="40%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Notes</th>
        <th width="20%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Actions</th>
      </tr>
      ${followUpTasks.map(task => {
        // Try multiple patterns to find email links
        let emailLink = null;
        
        // Pattern 1: Standard "Email Link: URL" format
        const emailLinkMatch = task.notes ? task.notes.match(/Email Link: (https:\/\/mail\.google\.com\/[^\s]+)/) : null;
        if (emailLinkMatch) {
          emailLink = emailLinkMatch[1];
        }
        
        // Pattern 2: Any URL in the notes
        if (!emailLink && task.notes) {
          const urlMatch = task.notes.match(/(https?:\/\/[^\s]+)/);
          if (urlMatch && urlMatch[1].includes('mail.google.com')) {
            emailLink = urlMatch[1];
          }
        }
        
        // Pattern 3: Check if there's a thread ID in the notes
        if (!emailLink && task.notes) {
          const threadMatch = task.notes.match(/Thread: ([a-zA-Z0-9]+)/);
          if (threadMatch) {
            emailLink = `https://mail.google.com/mail/u/0/#all/${threadMatch[1]}`;
          }
        }
        
        // Clean up notes by removing the email link
        let cleanNotes = task.notes || '';
        if (emailLink) {
          cleanNotes = cleanNotes
            .replace(/Email Link: https:\/\/mail\.google\.com\/[^\s]+/, '')
            .replace(/Thread: [a-zA-Z0-9]+/, '')
            .replace(/(https?:\/\/[^\s]+)/, '')
            .trim();
        }
        
        // Create a compose link for follow-up emails
        const subject = `Follow-up: ${task.name}`;
        const body = `Following up regarding: ${task.name}\n\n`;
        const composeLink = `https://mail.google.com/mail/u/0/?view=cm&fs=1&tf=1&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
        
        return `
          <tr style="background-color:${task.priority === 'P1' ? '#ffebee' : task.priority === 'P2' ? '#fff8e1' : task.priority === 'P3' ? '#e8f5e9' : '#e3f2fd'};">
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top; overflow:hidden; text-overflow:ellipsis;">${task.name}</td>
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top;">${cleanNotes}</td>
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top; white-space:nowrap;">
              ${emailLink ? 
                `<a href="${emailLink}" target="_blank" style="color:#3498db; text-decoration:none;">📧 Open</a><br>` : 
                ''}
              <a href="${composeLink}" target="_blank" style="color:#3498db; text-decoration:none;">✉️ Compose</a>
            </td>
          </tr>
        `;
      }).join('')}
    </table>
  `;
}

/**
 * Generate scheduled tasks table with scheduled time from spreadsheet
 */
function generateScheduledTasksHtml(scheduledTasks, hasCalendarAccess) {
  if (scheduledTasks.length === 0) {
    return '<p class="empty-message">📅 No scheduled tasks.</p>';
  }
  
  // Sort scheduled tasks by scheduled time
  const sortedTasks = [...scheduledTasks];
  
  // Try to get scheduled time from the spreadsheet first
  try {
    // Get the Tasks sheet to check for a Scheduled Time column
    const sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
    if (sheet) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // Look for Scheduled Time column
      const scheduledTimeColIndex = headers.findIndex(header => 
        String(header).toLowerCase() === 'scheduled time'
      );
      
      if (scheduledTimeColIndex >= 0) {
        console.log('Found Scheduled Time column at index:', scheduledTimeColIndex);
        
        // Get all scheduled time values
        const rows = sheet.getDataRange().getValues();
        
        // Add scheduled time to tasks from the spreadsheet
        sortedTasks.forEach(task => {
          if (task.row > 1 && task.row <= rows.length) {
            const scheduledTimeValue = rows[task.row - 1][scheduledTimeColIndex];
            if (scheduledTimeValue && scheduledTimeValue instanceof Date) {
              task.scheduledTime = scheduledTimeValue;
            }
          }
        });
      }
    }
  } catch (error) {
    console.error('Error getting scheduled times from spreadsheet:', error);
  }
  
  // For tasks without scheduled time, use a default time
  sortedTasks.forEach(task => {
    if (!task.scheduledTime) {
      // Default to today at 10:00 AM
      const defaultTime = new Date();
      defaultTime.setHours(10, 0, 0, 0);
      task.scheduledTime = defaultTime;
      task.estimatedTime = true; // Mark as estimated
    }
  });
  
  // Sort by scheduled time
  sortedTasks.sort((a, b) => {
    return a.scheduledTime.getTime() - b.scheduledTime.getTime();
  });
  
  // Generate the HTML table
  return `
    <table cellspacing="0" cellpadding="4" border="1" style="width:100%; border-collapse:collapse; font-size:13px; border:1px solid #ddd;">
      <tr>
        <th width="35%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Task</th>
        <th width="10%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Priority</th>
        <th width="10%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Time</th>
        <th width="20%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Scheduled</th>
        <th width="25%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Notes</th>
      </tr>
      ${sortedTasks.map(task => {
        // Format the scheduled time with emoji based on how soon it is
        let timeDisplay = 'Not set';
        let timeEmoji = '';
        
        if (task.scheduledTime) {
          const now = new Date();
          const timeDiff = task.scheduledTime.getTime() - now.getTime();
          const daysDiff = timeDiff / (1000 * 60 * 60 * 24);
          
          // Choose emoji based on how soon the task is scheduled
          if (daysDiff < 0) timeEmoji = '⏱️'; // Past
          else if (daysDiff < 1) timeEmoji = '🔥'; // Today
          else if (daysDiff < 2) timeEmoji = '⏰'; // Tomorrow
          else timeEmoji = '📅'; // Future
          
          // Format date in a more compact way
          const options = { 
            month: 'short', 
            day: 'numeric', 
            hour: 'numeric', 
            minute: '2-digit'
          };
          
          timeDisplay = `${timeEmoji} ${task.scheduledTime.toLocaleDateString(undefined, options)}`;
          
          // Add indicator if time is estimated
          if (task.estimatedTime) {
            timeDisplay += ' (est.)';
          }
        }
        
        // Choose priority emoji
        let priorityEmoji = '';
        switch(task.priority) {
          case 'P1': priorityEmoji = '🔴'; break;
          case 'P2': priorityEmoji = '🟠'; break;
          case 'P3': priorityEmoji = '🟢'; break;
          default: priorityEmoji = '⚪'; break;
        }
        
        // Clean up notes - remove email links to save space
        let cleanNotes = task.notes || '';
        cleanNotes = cleanNotes.replace(/Email Link: https:\/\/mail\.google\.com\/[^\s]+/g, '').trim();
        
        // Extract email links for the actions column
        const emailLinkMatch = task.notes ? task.notes.match(/Email Link: (https:\/\/mail\.google\.com\/[^\s]+)/) : null;
        const emailLink = emailLinkMatch ? emailLinkMatch[1] : null;
        
        return `
          <tr style="background-color:${task.priority === 'P1' ? '#ffebee' : task.priority === 'P2' ? '#fff8e1' : task.priority === 'P3' ? '#e8f5e9' : '#e3f2fd'};">
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top; overflow:hidden; text-overflow:ellipsis;">${task.name}</td>
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top; white-space:nowrap;">${priorityEmoji} ${task.priority}</td>
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top; white-space:nowrap;">${task.timeBlock}m</td>
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top; white-space:nowrap;">${timeDisplay}</td>
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top;">
              ${cleanNotes}
              ${emailLink ? `<div style="font-size:12px; margin-top:4px;"><a href="${emailLink}" target="_blank" style="color:#3498db; text-decoration:none;">📧 Email</a></div>` : ''}
            </td>
          </tr>
        `;
      }).join('')}
    </table>
  `;
}

/**
 * Generate error tasks table
 */
function generateErrorTasksHtml(errorTasks) {
  if (errorTasks.length === 0) {
    return '<p class="empty-message">✅ No tasks with errors.</p>';
  }
  
  return `
    <table cellspacing="0" cellpadding="4" border="1" style="width:100%; border-collapse:collapse; font-size:13px; border:1px solid #ddd;">
      <tr>
        <th width="40%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Task</th>
        <th width="15%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Priority</th>
        <th width="15%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Status</th>
        <th width="30%" style="background-color:#f8f9fa; text-align:left; padding:5px; border:1px solid #ddd;">Notes</th>
      </tr>
      ${errorTasks.map(task => {
        return `
          <tr style="background-color:#ffebee;">
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top; overflow:hidden; text-overflow:ellipsis;">${task.name}</td>
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top; white-space:nowrap;">${task.priority}</td>
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top; white-space:nowrap;">❌ ${task.status}</td>
            <td style="padding:4px; border:1px solid #ddd; vertical-align:top;">${task.notes || ''}</td>
          </tr>
        `;
      }).join('')}
    </table>
  `;
}

/**
 * Set up a daily trigger for the digest email
 * @param {string} timeString - Time to send digest in 24-hour format (e.g., "07:00")
 */
function setupDailyDigestTrigger(timeString = "07:00") {
  // Remove any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyTaskDigest') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Parse the time string
  const [hours, minutes] = timeString.split(':').map(num => parseInt(num, 10));
  
  // Create a new trigger
  ScriptApp.newTrigger('sendDailyTaskDigest')
    .timeBased()
    .atHour(hours)
    .nearMinute(minutes)
    .everyDays(1)
    .create();
  
  console.log(`Daily digest trigger set for ${timeString}`);
  
  // Store the setting
  PropertiesService.getScriptProperties().setProperty('dailyDigestTime', timeString);
  
  return true;
}

/**
 * Enable daily digest emails
 */
function enableDailyDigest() {
  const ui = SpreadsheetApp.getUi();
  
  // Get current time setting or use default
  const currentTime = PropertiesService.getScriptProperties().getProperty('dailyDigestTime') || "07:00";
  
  // Prompt for time
  const response = ui.prompt(
    'Enable Daily Digest',
    `Enter the time to send the daily digest email (24-hour format, e.g., "07:00").\nCurrent setting: ${currentTime}`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const timeString = response.getResponseText().trim() || currentTime;
    
    // Validate time format
    if (/^([01]?[0-9]|2[0-3]):([0-5][0-9])$/.test(timeString)) {
      // Set up the trigger
      setupDailyDigestTrigger(timeString);
      
      // Enable the feature
      PropertiesService.getScriptProperties().setProperty('dailyDigestEnabled', 'true');
      
      ui.alert('Success', `✅ Daily digest enabled and will be sent at ${timeString} every day.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', '❌ Invalid time format. Please use 24-hour format (e.g., "07:00").', ui.ButtonSet.OK);
    }
  }
}

/**
 * Disable daily digest emails
 */
function disableDailyDigest() {
  const ui = SpreadsheetApp.getUi();
  
  // Confirm with user
  const response = ui.alert(
    'Disable Daily Digest',
    'Are you sure you want to disable the daily digest emails?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // Remove any existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'sendDailyTaskDigest') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Disable the feature
    PropertiesService.getScriptProperties().setProperty('dailyDigestEnabled', 'false');
    
    ui.alert('Success', '✅ Daily digest emails have been disabled.', ui.ButtonSet.OK);
  }
}

/**
 * Test the daily digest by sending it immediately
 */
function testDailyDigest() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Temporarily enable the feature for testing
    const wasEnabled = PropertiesService.getScriptProperties().getProperty('dailyDigestEnabled') === 'true';
    PropertiesService.getScriptProperties().setProperty('dailyDigestEnabled', 'true');
    
    // Send the digest
    const result = sendDailyTaskDigest();
    
    // Restore previous setting
    if (!wasEnabled) {
      PropertiesService.getScriptProperties().setProperty('dailyDigestEnabled', 'false');
    }
    
    if (result) {
      ui.alert('Success', '✅ Daily digest email sent successfully. Check your inbox.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', '❌ Failed to send daily digest. Check the logs for details.', ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('Error testing daily digest:', error);
    ui.alert('Error', '❌ Failed to send daily digest: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Check the status of the daily digest feature
 */
function checkDailyDigestStatus() {
  const ui = SpreadsheetApp.getUi();
  
  const isEnabled = PropertiesService.getScriptProperties().getProperty('dailyDigestEnabled') === 'true';
  const timeString = PropertiesService.getScriptProperties().getProperty('dailyDigestTime') || "Not set";
  
  // Check if trigger exists
  let triggerExists = false;
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyTaskDigest') {
      triggerExists = true;
    }
  });
  
  const status = `📊 Daily Digest Status:
  
Enabled: ${isEnabled ? '✅ Yes' : '❌ No'}
Scheduled Time: ${timeString}
Trigger Active: ${triggerExists ? '✅ Yes' : '❌ No'}

${!isEnabled && triggerExists ? '⚠️ Warning: Trigger exists but feature is disabled.' : ''}
${isEnabled && !triggerExists ? '⚠️ Warning: Feature is enabled but trigger is missing.' : ''}`;
  
  ui.alert('Daily Digest Status', status, ui.ButtonSet.OK);
}

/**
 * Debug function to check follow-up tasks in the system
 */
function debugFollowUpTasks() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Get all tasks
    const allTasks = sheetManager.getTasks();
    
    // Log all tasks with their priorities
    console.log('All tasks:');
    allTasks.forEach(task => {
      console.log(`Task: ${task.name}, Priority: "${task.priority}", Status: "${task.status}"`);
    });
    
    // Try different ways to find follow-up tasks
    const exactMatch = allTasks.filter(task => task.priority === 'Follow-Up');
    const lowercaseMatch = allTasks.filter(task => String(task.priority || '').toLowerCase() === 'follow-up');
    const containsMatch = allTasks.filter(task => String(task.priority || '').toLowerCase().includes('follow'));
    
    // Show results
    const message = `Follow-up task detection:
    
Exact match "Follow-Up": ${exactMatch.length} tasks
Lowercase match "follow-up": ${lowercaseMatch.length} tasks
Contains "follow": ${containsMatch.length} tasks

Check the logs for more details.`;
    
    ui.alert('Debug Results', message, ui.ButtonSet.OK);
    
    // If we found any tasks with the contains method, log them
    if (containsMatch.length > 0) {
      console.log('Tasks containing "follow" in priority:');
      containsMatch.forEach(task => {
        console.log(`- ${task.name} (Priority: "${task.priority}")`);
      });
    }
    
    return {
      exactMatch,
      lowercaseMatch,
      containsMatch
    };
  } catch (error) {
    console.error('Error debugging follow-up tasks:', error);
    ui.alert('Error', 'Failed to debug follow-up tasks: ' + error.message, ui.ButtonSet.OK);
    return null;
  }
}

/**
 * Utility to add email links to follow-up tasks
 */
function addEmailLinkToTask() {
  const ui = SpreadsheetApp.getUi();
  
  // Get the active sheet and selected cell
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'Tasks') {
    ui.alert('Error', 'Please select a task in the Tasks sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  
  // Skip header row
  if (row === 1) {
    ui.alert('Error', 'Please select a task row, not the header.', ui.ButtonSet.OK);
    return;
  }
  
  // Get the task
  const task = sheetManager.getTaskByRow(row);
  if (!task || !task.name) {
    ui.alert('Error', 'Could not find a valid task in the selected row.', ui.ButtonSet.OK);
    return;
  }
  
  // Prompt for the email link or thread ID
  const response = ui.prompt(
    'Add Email Link',
    `Enter the Gmail thread ID or full URL for task "${task.name}":`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText().trim();
    if (!input) {
      ui.alert('Error', 'No link or ID provided.', ui.ButtonSet.OK);
      return;
    }
    
    let emailLink;
    // Check if it's a full URL or just a thread ID
    if (input.startsWith('http')) {
      emailLink = input;
    } else {
      emailLink = `https://mail.google.com/mail/u/0/#all/${input}`;
    }
    
    // Update the notes field
    const notesColumn = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('Notes') + 1;
    if (notesColumn < 1) {
      ui.alert('Error', 'Could not find Notes column.', ui.ButtonSet.OK);
      return;
    }
    
    const currentNotes = sheet.getRange(row, notesColumn).getValue() || '';
    const updatedNotes = currentNotes + (currentNotes ? '\n' : '') + `Email Link: ${emailLink}`;
    
    sheet.getRange(row, notesColumn).setValue(updatedNotes);
    ui.alert('Success', `Email link added to task "${task.name}".`, ui.ButtonSet.OK);
  }
}

/**
 * Fix follow-up tasks with incorrect capitalization
 * This is called from the menu
 */
function fixFollowUpTasks() {
  // Call our improved standardization function
  standardizeFollowUpPriorities();
} 