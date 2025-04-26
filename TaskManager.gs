/**
 * Task Manager coordinates between sheets and calendar
 */
class TaskManager {
  constructor() {
    // Don't initialize settings in constructor
    this.settings = null;
    this.scheduledEvents = []; // Add this to track events we schedule
  }

  /**
   * Initialize settings
   */
  initialize() {
    // Load settings
    this.settings = initializeSettings();
    
    // Try to initialize calendar manager, but continue if it fails
    try {
      this.calendarManager = calendarManager;
      this.hasCalendarAccess = this.calendarManager && this.calendarManager.isCalendarAvailable();
      console.log('Calendar access available:', this.hasCalendarAccess);
    } catch (error) {
      console.error('Error initializing calendar manager:', error);
      this.hasCalendarAccess = false;
    }
    
    console.log('Task manager initialized with settings:', JSON.stringify(this.settings));
  }

  /**
   * Find a suitable time slot for a task
   * @param {Object} task - Task object
   * @returns {Object|null} Available slot or null if none found
   */
  findSlotForTask(task) {
    // This is now just a wrapper that uses the processTask method
    // to find a slot without actually scheduling the task
    const calendar = CalendarApp.getDefaultCalendar();
    const now = new Date();
    const maxDaysToCheck = 7; // Number of days to try scheduling
    const duration = task.timeBlock || 30;
    
    // Add priority delay
    let priorityDelay = 0;
    if (task.priority === 'P1') {
      priorityDelay = 30; // 30 minutes delay for P1
    } else if (task.priority === 'P2') {
      priorityDelay = 60; // 1 hour delay for P2
    } else if (task.priority === 'P3') {
      priorityDelay = 120; // 2 hours delay for P3
    }
    
    // Apply priority delay to now
    const searchStart = new Date(now);
    searchStart.setMinutes(searchStart.getMinutes() + priorityDelay);
    
    for (let i = 0; i < maxDaysToCheck; i++) {
      const dayToCheck = new Date(searchStart);
      dayToCheck.setDate(searchStart.getDate() + i);
      
      const weekday = dayToCheck.toLocaleDateString('en-US', { weekday: 'long' });
      const daySchedule = this.settings.workHours[weekday];
      if (!daySchedule) {
        console.log(`No work hours defined for ${weekday}, skipping day`);
        continue; // Skip if no work hours for this day
      }
      
      console.log(`Checking ${weekday} for available slots`);
      
      for (const block of daySchedule) {
        const startTimeParts = block.start.split(':');
        const endTimeParts = block.end.split(':');
        
        const start = new Date(dayToCheck);
        start.setHours(+startTimeParts[0], +startTimeParts[1], 0, 0);
        
        const end = new Date(dayToCheck);
        end.setHours(+endTimeParts[0], +endTimeParts[1], 0, 0);
        
        // Skip if this block is in the past
        if (end <= now) {
          console.log(`Work block ${block.start}-${block.end} is in the past, skipping`);
          continue;
        }
        
        // If this is today and the block start is in the past, adjust start time
        if (dayToCheck.toDateString() === now.toDateString() && start < now) {
          start.setTime(now.getTime());
          // Round up to next 15-minute mark
          const minutes = start.getMinutes();
          if (minutes > 45) {
            start.setHours(start.getHours() + 1, 0, 0, 0);
          } else if (minutes > 30) {
            start.setMinutes(45, 0, 0);
          } else if (minutes > 15) {
            start.setMinutes(30, 0, 0);
          } else if (minutes > 0) {
            start.setMinutes(15, 0, 0);
          }
        }
        
        // Try slots in 15-min increments
        while (start.getTime() + duration * 60000 <= end.getTime()) {
          const slotEnd = new Date(start.getTime() + duration * 60000);
          const potentialEvents = calendar.getEvents(start, slotEnd);

          // Filter events: only count accepted, busy events as conflicts
          const conflictingEvents = potentialEvents.filter(event => {
            const transparency = event.getTransparency();
            const myStatus = event.getMyStatus();
            // Log details for EVERY potential event, including type of transparency
            console.log(`Filtering event: \"${event.getTitle()}\", Status: ${myStatus}, Transparency Value: ${transparency}`);

            const isAccepted = !myStatus || myStatus === CalendarApp.GuestStatus.YES || myStatus === CalendarApp.GuestStatus.OWNER;
            // Compare with both enum and potential string value
            const isBusy = transparency === CalendarApp.Visibility.DEFAULT || transparency === CalendarApp.Visibility.OPAQUE || String(transparency).toUpperCase() === 'OPAQUE';

            // Log filter evaluation results
            console.log(` -> isAccepted: ${isAccepted}, isBusy: ${isBusy} (Requires BOTH true to be conflict)`); 
            return isAccepted && isBusy;
          });

          if (conflictingEvents.length === 0) {
            console.log(`Found available slot: ${start.toLocaleString()} - ${slotEnd.toLocaleString()}`);
            return {
              start: new Date(start),
              end: new Date(slotEnd)
            };
          } else {
             // Log conflicting events for debugging
             conflictingEvents.forEach(ev => {
               console.log(`Slot conflict found: ${ev.getTitle()} at ${ev.getStartTime().toLocaleString()}, Status: ${ev.getMyStatus()}, Transparency: ${ev.getTransparency()}`);
             });
          }

          start.setMinutes(start.getMinutes() + 15); // slide window forward
        }
      }
    }
    
    console.warn(`No available slots in next ${maxDaysToCheck} days for "${task.name}"`);
    return null;
  }

  /**
   * Schedule a task based on priority rules
   * @param {Object} task - Task object
   * @returns {Object} Scheduled task details
   */
  async scheduleTask(task) {
    if (!this.settings) {
      this.initialize();
    }

    // Set default time block if empty
    if (!task.timeBlock) {
      task.timeBlock = 30; // Default to 30 minutes
      console.log(`Setting default time block for task "${task.name}"`);
    }

    // Skip paused tasks and follow-ups (check both capitalization variants)
    const taskPriority = String(task.priority || '').toLowerCase();
    if (task.status === 'Pause' || 
        taskPriority === 'follow-up' || 
        taskPriority === 'follow up' ||
        taskPriority === 'followup') {
      console.log(`Skipping ${taskPriority.includes('follow') ? 'follow-up' : 'paused'} task:`, task.name);
      return task;
    }

    // Check if task is too long for any work period
    const maxSlotDuration = Math.max(...Object.values(this.settings.workHours)
      .flat()
      .map(period => {
        const start = calendarManager.parseTime(period.start);
        const end = calendarManager.parseTime(period.end);
        return (end.hours - start.hours) * 60 + (end.minutes - start.minutes);
      }));

    if (task.timeBlock > maxSlotDuration) {
      console.error(`Task "${task.name}" duration (${task.timeBlock} mins) exceeds maximum slot duration (${maxSlotDuration} mins)`);
      return task;
    }

    // Find suitable slot based on priority rules
    const slot = await this.findSlotForTask(task);

    if (!slot) {
      console.error('No slots found after extensive search for task:', task.name);
      return task;
    }

    try {
      const scheduledEvent = calendarManager.scheduleTask(task, slot.start);
      this.scheduledEvents.push(scheduledEvent); // Add to our tracking
      
      // Update sheet first
      sheetManager.updateTaskStatus(task.row, 'Scheduled');

      // Format time and date for email subject
      const timeStr = Utilities.formatDate(scheduledEvent.start, PropertiesService.getScriptProperties().getProperty('timeZone') || 'GMT', '@h:mm a');
      const dateStr = Utilities.formatDate(scheduledEvent.start, PropertiesService.getScriptProperties().getProperty('timeZone') || 'GMT', "'of' EEE (M/d)");
      
      // Send email notification with consistent format
      const userEmail = PropertiesService.getScriptProperties().getProperty('userEmail');
      if (userEmail) {
        MailApp.sendEmail({
          to: userEmail,
          subject: `Task Created: ${task.name} ${timeStr} ${dateStr}`,
          body: `Your task has been scheduled:

Task: ${task.name}
Time: ${scheduledEvent.start.toLocaleString()}
Priority: ${task.priority}
Duration: ${task.timeBlock} minutes

Notes: ${task.notes || 'None'}

This is an automated message from your Digital Assistant.`
        });
      }

      const newEvent = {
        title: scheduledEvent.title,
        start: scheduledEvent.start,
        end: scheduledEvent.end
      };

      // Verify no existing event has the exact same start time
      const conflictingEvent = this.scheduledEvents.find(event => {
        const eventStart = calendarManager.isCustomEvent(event) ? 
          event.start : event.getStartTime();
        return eventStart.getTime() === newEvent.start.getTime();
      });

      if (conflictingEvent) {
        console.log(`WARNING: Detected conflict - "${newEvent.title}" conflicts with "${
          calendarManager.isCustomEvent(conflictingEvent) ? 
          conflictingEvent.title : conflictingEvent.getTitle()
        }"`);
        // Handle conflict by moving to next slot, adjusting time, etc.
      } else {
        // No conflict, safe to add
        this.scheduledEvents.push(newEvent);
        console.log(`Added event "${scheduledEvent.title}" to tracking list, now tracking ${this.scheduledEvents.length} events`);
      }

      // After successfully scheduling the task, update the Scheduled Time in the spreadsheet
      const scheduledTime = scheduledEvent.getStartTime();
      updateTaskScheduledTime(task, scheduledTime);

      return scheduledEvent;
    } catch (error) {
      console.error('Error scheduling task', task.name + ':', error);
      throw error;
    }
  }

  /**
   * Process all pending tasks
   * @returns {Array} Array of scheduling results
   */
  processPendingTasks() {
    if (!this.settings) this.initialize();
    
    // Get all pending tasks
    const pendingTasks = sheetManager.getPendingTasks();
    console.log(`Found ${pendingTasks.length} pending tasks`);
    
    // Process tasks in priority order
    const priorityOrder = ['P1', 'P2', 'P3'];
    const results = [];
    
    // Process each priority level
    for (const priority of priorityOrder) {
      const tasksWithPriority = pendingTasks.filter(task => task.priority === priority);
      console.log(`Processing ${tasksWithPriority.length} ${priority} tasks`);
      
      // Process each task with this priority
      for (const task of tasksWithPriority) {
        console.log(`Processing ${priority} task: ${task.name}`);
        const result = this.processTask(task);
        results.push({
          task: task.name,
          priority: task.priority,
          success: result.success,
          message: result.message,
          events: result.events
        });
      }
    }
    
    // Process any remaining tasks (those without a priority)
    const tasksWithoutPriority = pendingTasks.filter(task => !priorityOrder.includes(task.priority));
    console.log(`Processing ${tasksWithoutPriority.length} tasks without priority`);
    
    for (const task of tasksWithoutPriority) {
      console.log(`Processing task without priority: ${task.name}`);
      const result = this.processTask(task);
      results.push({
        task: task.name,
        priority: task.priority || 'None',
        success: result.success,
        message: result.message,
        events: result.events
      });
    }
    
    console.log(`Processed ${results.length} tasks total`);
    return results;
  }

  /**
   * Schedule a task with priority rules
   * @param {Object} task - Task to schedule
   * @param {number} hoursDelay - Hours to delay scheduling
   * @param {Array} newlyScheduledEvents - Array to store newly scheduled events
   * @param {Array} existingEvents - Existing events to consider
   * @returns {Object} Scheduling result
   */
  async scheduleTaskWithPriorityRules(task, hoursDelay, newlyScheduledEvents, existingEvents) {
    try {
      console.log(`Priority ${task.priority}: Scheduling to start after ${hoursDelay} hours (${new Date(Date.now() + hoursDelay * 60 * 60 * 1000).toLocaleString()})`);
      
      // Find a suitable slot
      const startDate = new Date();
      startDate.setHours(startDate.getHours() + hoursDelay);
      
      console.log(`Checking for conflicts with ${existingEvents.length} calendar events + ${newlyScheduledEvents.length} newly scheduled events`);
      
      // Combine existing and newly scheduled events for conflict checking
      const allEvents = [...existingEvents, ...newlyScheduledEvents];
      
      // Find available slot
      const slots = await calendarManager.findAvailableSlots(
        task.timeBlock || 30,
        startDate,
        this.settings.workHours,
        allEvents
      );
      
      if (!slots || slots.length === 0) {
        console.error(`No available slots found for task "${task.name}"`);
        return { success: false, message: 'No available slots found' };
      }
      
      console.log(`Scheduled task "${task.name}" for ${slots[0].start}`);
      
      // Create calendar event
      const event = calendarManager.createCalendarEvent(task, slots[0].start);
      
      // Add to tracking list
      const newEvent = {
        title: `Auto-Scheduled: ${task.name}`,
        start: slots[0].start,
        end: new Date(slots[0].start.getTime() + (task.timeBlock || 30) * 60 * 1000)
      };
      
      newlyScheduledEvents.push(newEvent);
      console.log(`Added event "Auto-Scheduled: ${task.name}" to tracking list, now tracking ${newlyScheduledEvents.length} events`);
      
      return { success: true, event: newEvent };
    } catch (error) {
      console.error(`Error scheduling task "${task.name}":`, error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Schedule multiple short tasks optimized to be back-to-back
   */
  async scheduleShortTasksOptimized(tasks, hoursDelay, newlyScheduledEvents, allEvents) {
    if (tasks.length === 0) return;
    
    // Calculate total duration needed
    const totalDuration = tasks.reduce((sum, task) => sum + (task.timeBlock || 15), 0);
    console.log(`Optimizing ${tasks.length} short tasks with total duration ${totalDuration} minutes`);
    
    // Create a date with the appropriate delay
    const startDate = new Date();
    startDate.setHours(startDate.getHours() + hoursDelay);
    
    // Set up search parameters
    let searchDate = new Date(startDate);
    let scheduled = false;
    let attempts = 0;
    const MAX_DAYS_TO_TRY = 30;
    
    // Keep trying until scheduled or max attempts reached
    while (!scheduled && attempts < MAX_DAYS_TO_TRY) {
      // Get existing calendar events for this day
      const existingEvents = calendarManager.getEventsForDay(searchDate);
      const combinedEvents = [...existingEvents, ...newlyScheduledEvents];
      
      console.log(`Checking for conflicts with ${existingEvents.length} calendar events + ${newlyScheduledEvents.length} newly scheduled events`);
      
      // Find available slots for the total duration
      const slots = await calendarManager.findAvailableSlots(
        totalDuration,
        searchDate,
        this.settings.workHours,
        combinedEvents
      );
      
      if (slots.length > 0) {
        // We found a slot big enough for all tasks
        let currentStart = slots[0].start;
        
        // Schedule each task back-to-back
        for (const task of tasks) {
          const duration = task.timeBlock || 15;
          const taskEnd = new Date(currentStart.getTime() + duration * 60000);
          
          // Schedule in calendar
          const scheduledEvent = calendarManager.scheduleTask(task, currentStart, taskEnd);
          this.markTaskScheduled(task.row);
          
          // Add to tracking
          const newEvent = {
            title: scheduledEvent.title,
            start: scheduledEvent.start,
            end: scheduledEvent.end
          };
          
          newlyScheduledEvents.push(newEvent);
          console.log(`Scheduled optimized task "${task.name}" for ${currentStart.toLocaleString()} to ${taskEnd.toLocaleString()}`);
          
          // Set start time for next task
          currentStart = taskEnd;
        }
        
        scheduled = true;
        console.log(`Successfully scheduled ${tasks.length} tasks back-to-back`);
      } else {
        // Move to next day
        searchDate.setDate(searchDate.getDate() + 1);
        searchDate.setHours(0, 0, 0, 0);
        attempts++;
        console.log(`No slots found for optimized tasks, trying next day`);
      }
    }
    
    if (!scheduled) {
      console.log(`Could not schedule optimized tasks within ${MAX_DAYS_TO_TRY} days, falling back to individual scheduling`);
      // Fall back to scheduling individually
      for (const task of tasks) {
        await this.scheduleTaskWithPriorityRules(task, hoursDelay, newlyScheduledEvents, newlyScheduledEvents);
      }
    }
  }

  /**
   * Get all pending tasks
   * @returns {Array} Array of pending tasks
   */
  getPendingTasks() {
    // Get all tasks from the sheet
    const allTasks = sheetManager.getTasks();
    
    // Filter for pending tasks
    const pendingTasks = allTasks.filter(task => 
      task.status === 'Pending' || !task.status
    );
    
    console.log(`Found ${pendingTasks.length} pending tasks out of ${allTasks.length} total tasks`);
    return pendingTasks;
  }

  /**
   * Mark a task as scheduled in the sheet
   * @param {number} row - Row number of the task
   */
  markTaskScheduled(row) {
    sheetManager.updateTaskStatus(row, 'Scheduled');
  }

  /**
   * Process a single task
   * @param {Object} task - Task to process
   * @param {string} [fallbackTaskName] - Fallback name if task.name is empty
   * @returns {Object} Processing result
   */
  processTask(task, fallbackTaskName) {
    if (!this.settings) this.initialize();

    // --- Robust check for Follow-up and Paused tasks --- 
    const taskPriority = task.priority || 'P1'; // Default blank priority to P1 for processing logic
    const taskPriorityLower = String(taskPriority).toLowerCase().trim();
    const taskStatusLower = String(task.status || '').toLowerCase().trim();

    // --- Skip scheduling if STATUS is 'Follow-up' ---
    if (taskStatusLower === 'follow-up') {
      console.log(`Skipping task due to 'Follow-up' status: "${task.name}"`);
      return {
        success: true, 
        message: 'Task skipped due to Follow-up status',
        events: [] // Ensure events array is returned
      };
    }
    // --- End Skip Check ---

    console.log(`Processing task: ${task.name} (Row: ${task.row}), Priority: ${taskPriority}, Status: ${task.status}`);

    // Skip paused tasks
    if (taskStatusLower === 'pause') {
      console.log('Skipping paused task:', task.name);
      return {
        success: true, 
        message: 'Paused task skipped',
        events: [] // Ensure events array is returned
      };
    }
    // --- End of checks --- 
    
    const calendar = CalendarApp.getDefaultCalendar();
    const now = new Date();
    const today = new Date(now); // Use a stable 'today' for offset calculation
    today.setHours(0, 0, 0, 0); 

    // --- Calculate Target Start Date based on Priority ---
    let workingDaysOffset = 0;
    if (taskPriority === 'P2') {
      workingDaysOffset = 2;
    } else if (taskPriority === 'P3') {
      workingDaysOffset = 3;
    } // P1 has 0 offset

    let searchStartDate = new Date(now); // Start search from current time by default (for P1)
    if (workingDaysOffset > 0) {
      const targetWorkingDay = calculateWorkingDayOffset(today, workingDaysOffset, this.settings.workHours);
      // --- FIX: Start searching the day AFTER the target working day ---
      searchStartDate = new Date(targetWorkingDay); 
      searchStartDate.setDate(targetWorkingDay.getDate() + 1); 
      // --- End FIX ---
      searchStartDate.setHours(0, 0, 0, 0); // Ensure we start at the beginning of that day
      console.log(`Priority ${taskPriority}: Target is AFTER ${workingDaysOffset} working days. Starting search on: ${searchStartDate.toLocaleString()}`);
    } else {
       console.log(`Priority ${taskPriority}: Starting search immediately.`);
       // Ensure P1 search starts effectively from 'now', not midnight, if we adjusted searchStartDate
       searchStartDate = new Date(now); 
    }
    // --- End Target Date Calculation ---

    const maxDaysToCheck = 7; 
    const duration = task.timeBlock || 30;
    
    // --- IMPORTANT: Adjust the initial searchStartDate for the loop logic ---
    // The loop adds 'i' days. To start searching ON searchStartDate, we need
    // to subtract one day here so that when i=0, dayToCheck is correct.
    let loopBaseDate = new Date(searchStartDate);
    loopBaseDate.setDate(loopBaseDate.getDate() - 1); 
    // --- End adjustment ---

    console.log(`Starting search loop, effectively beginning checks on: ${searchStartDate.toLocaleString()}`);
    
    for (let i = 0; i < maxDaysToCheck; i++) {
       // --- Calculate dayToCheck based on loopBaseDate ---
      const dayToCheck = new Date(loopBaseDate);
      dayToCheck.setDate(loopBaseDate.getDate() + i + 1); // Add i+1 days to the base
      // --- End calculation ---
      
      const dayToCheckDateString = dayToCheck.toDateString();
      
      // Check against 'today' (start of today) for relevance
      const isTodayOrFuture = dayToCheck.getTime() >= today.getTime(); 

      // This check might not be strictly needed anymore with the corrected start date,
      // but keep as a safety measure.
      if (!isTodayOrFuture && i > 0) { // Allow the calculated start date even if it's somehow "yesterday" relative to exact 'now'
         console.log(`Skipping past day: ${dayToCheckDateString}`);
         continue; 
      }

      const weekday = dayToCheck.toLocaleDateString('en-US', { weekday: 'long' });
      const daySchedule = this.settings.workHours[weekday];
      if (!daySchedule) {
        console.log(`No work hours defined for ${weekday} (${dayToCheckDateString}), skipping day`);
        continue; 
      }
      
      console.log(`Checking ${weekday} (${dayToCheckDateString}) for available slots`);
      
      let foundSlotInThisDay = false;
      
      for (const block of daySchedule) {
        const startTimeParts = block.start.split(':');
        const endTimeParts = block.end.split(':');
        
        const blockStartDateTime = new Date(dayToCheck);
        blockStartDateTime.setHours(+startTimeParts[0], +startTimeParts[1], 0, 0);
        
        const blockEndDateTime = new Date(dayToCheck);
        blockEndDateTime.setHours(+endTimeParts[0], +endTimeParts[1], 0, 0);
        
        // Determine the actual start time for searching within this block
        let currentSearchTime = new Date(blockStartDateTime);

        // If we are checking the *initial* search day (i=0), 
        // and this block starts before the overall searchStartDate,
        // adjust the search start to the later of the two.
        // Also, ensure we don't start searching in the past relative to 'now'.
        if (i === 0 && currentSearchTime < searchStartDate) {
           currentSearchTime = new Date(searchStartDate);
        }
         // Ensure we don't start searching before 'now' if it's today
        if (currentSearchTime < now) {
           currentSearchTime = new Date(now);
           // Optional: Round up to next 15 min? Keep it simple for now.
        }

        // If the adjusted search time is now after the block ends, skip block
        if (currentSearchTime >= blockEndDateTime) {
            console.log(`Adjusted search start time (${currentSearchTime.toLocaleTimeString()}) is after block end (${blockEndDateTime.toLocaleTimeString()}), skipping block.`);
            continue;
        }
        
        console.log(`Checking block ${block.start}-${block.end} on ${weekday}, starting scan from ${currentSearchTime.toLocaleTimeString()}`);
        
        // Try slots in 15-min increments
        while (currentSearchTime.getTime() + duration * 60000 <= blockEndDateTime.getTime()) {
          const slotStart = new Date(currentSearchTime); // Potential start time
          const slotEnd = new Date(slotStart.getTime() + duration * 60000);
          let conflictingEvents = calendar.getEvents(slotStart, slotEnd);

          // Filter events: only count accepted, busy events as conflicts
          conflictingEvents = conflictingEvents.filter(event => {
            const transparency = event.getTransparency();
            const myStatus = event.getMyStatus();
            console.log(`Filtering event: "${event.getTitle()}", Status: ${myStatus}, Transparency Value: ${transparency}`);
            const isAccepted = !myStatus || myStatus === CalendarApp.GuestStatus.YES || myStatus === CalendarApp.GuestStatus.OWNER;
            const isBusy = transparency === CalendarApp.Visibility.DEFAULT || transparency === CalendarApp.Visibility.OPAQUE || String(transparency).toUpperCase() === 'OPAQUE';
            console.log(` -> isAccepted: ${isAccepted}, isBusy: ${isBusy} (Requires BOTH true to be conflict)`); 
            return isAccepted && isBusy;
          });

          let bumpedP3 = false; // Flag to indicate if we bumped a P3

          // --- P1 vs Auto-Scheduled P3 Bumping Logic ---
          if (taskPriority === 'P1' && conflictingEvents.length === 1) {
             const conflictingEvent = conflictingEvents[0];
             const eventTitle = conflictingEvent.getTitle();
             const eventDesc = conflictingEvent.getDescription();
             
             console.log(`P1 Conflict Check: Found 1 conflicting event: "${eventTitle}"`);

             // Check if it looks like one of our auto-scheduled events
             if (eventTitle.startsWith('Auto-Scheduled: ')) {
                const rowMatch = eventDesc.match(/Source Task Row: (\d+)/); // Escaped backslash for regex literal
                if (rowMatch && rowMatch[1]) {
                   const conflictingTaskRow = parseInt(rowMatch[1], 10);
                   console.log(`P1 Conflict Check: Conflicting event is Auto-Scheduled, corresponds to row ${conflictingTaskRow}`);
                   
                   // Fetch the original task details (requires sheet access)
                   const originalTask = sheetManager.getTaskByRow(conflictingTaskRow); // Assumes getTaskByRow exists and works
                   
                   if (originalTask && (originalTask.priority === 'P3')) {
                      console.log(`P1 BUMPING P3: Conflicting task in row ${conflictingTaskRow} is P3. Bumping.`);
                      // 1. Delete the P3 calendar event
                      try {
                         conflictingEvent.deleteEvent();
                         console.log(`P1 BUMPING P3: Deleted calendar event for P3 task "${originalTask.name || 'Unnamed'}".`);
                         // 2. Reset the P3 task status to Pending
                         const statusUpdated = sheetManager.updateTaskStatus(conflictingTaskRow, 'Pending');
                          if (statusUpdated) {
                            console.log(`P1 BUMPING P3: Reset status for P3 task in row ${conflictingTaskRow} to Pending.`);
                            conflictingEvents = []; // Clear conflicts as we just resolved it
                            bumpedP3 = true; // Mark that we bumped
                          } else {
                             console.error(`P1 BUMPING P3 FAILED: Could not reset status for P3 task in row ${conflictingTaskRow}. Aborting bump.`);
                             // Keep conflictingEvents array as is, P1 won't be scheduled here.
                          }
                      } catch (deleteError) {
                         console.error(`P1 BUMPING P3 FAILED: Error deleting calendar event: ${deleteError}. Aborting bump.`);
                         // Keep conflictingEvents array as is, P1 won't be scheduled here.
                      }
                   } else {
                     console.log(`P1 Conflict Check: Conflicting task is not P3 (Priority: ${originalTask ? originalTask.priority : 'Unknown'}). No bump.`);
                   }
                } else {
                  console.log(`P1 Conflict Check: Could not extract row number from event description: "${eventDesc}". No bump.`);
                }
             } else {
               console.log(`P1 Conflict Check: Conflicting event "${eventTitle}" is not an Auto-Scheduled event. No bump.`);
             }
          }
          // --- End Bumping Logic ---


          if (conflictingEvents.length === 0) {
            console.log(`Found available slot: ${slotStart.toLocaleString()} - ${slotEnd.toLocaleString()}${bumpedP3 ? ' (after bumping P3)' : ''}`);
            
            // Create the calendar event
            const event = calendar.createEvent(
              `Auto-Scheduled: ${task.name || fallbackTaskName || 'Unnamed Task'}`,
              slotStart, // Use the actual slot start time found
              slotEnd,   // Use the actual slot end time found
              {
                // --- ADD ROW INFO TO DESCRIPTION ---
                description: `Priority: ${taskPriority}\nTime Block: ${duration} minutes\nSource Task Row: ${task.row}\n\nAuto-scheduled by Digital Assistant.`, // Escaped newlines
                transparency: CalendarApp.EventTransparency.TRANSPARENT 
              }
            );
            
            event.setColor(CalendarApp.EventColor.ORANGE);
            
            // Update task status in sheet
            sheetManager.updateTaskStatus(task.row, 'Scheduled');

            // Update scheduled time column in sheet
            updateTaskScheduledTime(task, event.getStartTime()); 

            foundSlotInThisDay = true; // Mark that we found a slot today
            
            // Prepare result object for confirmation email etc.
             const resultEvent = {
               id: event.getId(),
               title: event.getTitle(),
               start: event.getStartTime(),
               end: event.getEndTime()
             };

            return {
              success: true,
              message: 'Task scheduled successfully',
              events: [resultEvent] // Return event details in an array
            };
          } else {
             // Log remaining conflicting events for debugging if bump didn't happen or failed
             if (!bumpedP3) {
               conflictingEvents.forEach(ev => {
                 console.log(`Slot conflict remains: ${ev.getTitle()} at ${ev.getStartTime().toLocaleString()}`);
               });
             }
          }

          // Slide window forward
          currentSearchTime.setMinutes(currentSearchTime.getMinutes() + 15); 
        } // End while loop for slots within block
      } // End for loop for blocks within day
      
       if (!foundSlotInThisDay) {
         console.log(`No available slots found for ${weekday}, moving to next day`);
       } else {
         // If we found a slot, we don't need to check subsequent days
         // This break is handled by the 'return' inside the loop
       }
    } // End for loop for days
    
    console.warn(`No available slots found within ${maxDaysToCheck} days (from target start) for "${task.name}"`);
    
    return {
      success: false,
      message: `No available slots found for task '${task.name}' within search window.`,
      events: []
    };
  } // End processTask method

  /**
   * Schedule a task on the calendar
   * @param {Object} task - Task to schedule
   * @param {Date} startTime - Start time
   * @param {Date} [endTime] - Optional end time (calculated from timeBlock if not provided)
   * @returns {Object} Scheduled event details
   */
  scheduleTask(task, startTime, endTime) {
    // Validate task object
    if (!task || !task.name) {
      console.error('Invalid task object for scheduling:', JSON.stringify(task));
      // Create a fallback task name
      task = task || {};
      task.name = task.name || 'Unnamed Task';
    }
    
    // Calculate end time if not provided
    const calculatedEndTime = endTime || new Date(startTime.getTime() + (task.timeBlock || 30) * 60000);
    
    // Create calendar event
    const event = this.calendar.createEvent(
      `Auto-Scheduled: ${task.name}`,
      startTime,
      calculatedEndTime,
      {
        description: `Priority: ${task.priority || 'Not set'}\nNotes: ${task.notes || 'None'}`,
        transparency: CalendarApp.EventTransparency.TRANSPARENT
      }
    );
    
    // Set event color to orange (6)
    event.setColor(CalendarApp.EventColor.ORANGE);
    
    console.log(`Scheduled task "${task.name}" for ${startTime}`);
    
    return {
      id: event.getId(),
      title: event.getTitle(),
      start: event.getStartTime(),
      end: event.getEndTime()
    };
  }

  /**
   * Clear auto-scheduled events for a specific task
   * @param {string} taskName - Name of the task
   * @returns {number} Number of events deleted
   */
  clearTaskEvents(taskName) {
    if (!this.isCalendarAvailable()) {
      console.warn('Calendar not available, skipping clearTaskEvents');
      return;
    }
    
    // Use a default name if taskName is undefined
    const searchName = taskName || 'Unnamed Task';
    
    const now = new Date();
    const startDate = new Date(now);
    startDate.setDate(startDate.getDate() - 1);
    
    const endDate = new Date(now);
    endDate.setDate(endDate.getDate() + 14);
    
    console.log(`Searching for events with title "Auto-Scheduled: ${searchName}"`);
    
    // ... rest of the function ...
  }
}

// Create global instance
const taskManager = new TaskManager();

/**
 * Update the Scheduled Time for a task in the spreadsheet
 * @param {Object} task - The task to update
 * @param {Date} scheduledTime - The scheduled time to record
 */
function updateTaskScheduledTime(task, scheduledTime) {
  try {
    console.log(`Updating scheduled time for task at row ${task ? task.row : 'undefined'} to ${scheduledTime}`);
    
    if (!task || !task.row || !scheduledTime) {
      console.error('Invalid parameters for updateTaskScheduledTime:', 
                   { task: task ? `row ${task.row}` : 'undefined', 
                     time: scheduledTime ? scheduledTime.toString() : 'undefined' });
      return;
    }
    
    const sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
    if (!sheet) {
      console.error('Tasks sheet not found');
      return;
    }
    
    // Get headers to find the Scheduled Time column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log(`Sheet headers: ${JSON.stringify(headers)}`);
    
    const scheduledTimeColIndex = headers.findIndex(header => 
      String(header).toLowerCase() === 'scheduled time'
    );
    
    console.log(`Scheduled Time column index: ${scheduledTimeColIndex}`);
    
    // If column exists, update the scheduled time
    if (scheduledTimeColIndex >= 0) {
      console.log(`Setting value at row ${task.row}, column ${scheduledTimeColIndex + 1} to ${scheduledTime}`);
      sheet.getRange(task.row, scheduledTimeColIndex + 1).setValue(scheduledTime);
      console.log(`Successfully updated Scheduled Time for task "${task.name || 'unnamed'}" to ${scheduledTime}`);
    } else {
      console.error('Scheduled Time column not found in Tasks sheet');
    }
  } catch (error) {
    console.error('Error updating scheduled time:', error);
    // Don't throw - this is a non-critical operation
  }
}

// Helper function to calculate a future date skipping non-working days
function calculateWorkingDayOffset(startDate, offsetDays, workHoursConfig) {
  let currentDate = new Date(startDate);
  let daysAdded = 0;
  
  // Ensure we start checking from the beginning of the day for consistency
  currentDate.setHours(0, 0, 0, 0); 

  while (daysAdded < offsetDays) {
    currentDate.setDate(currentDate.getDate() + 1); // Move to the next day
    const weekday = currentDate.toLocaleDateString('en-US', { weekday: 'long' });
    // Check if this weekday exists in the work hours config (meaning it's a working day)
    if (workHoursConfig[weekday]) {
      daysAdded++;
    }
  }
  // Return the date of the target working day (start of the day)
  return currentDate; 
} 