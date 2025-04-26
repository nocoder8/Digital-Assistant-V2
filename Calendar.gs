/**
 * Calendar operations for Digital Assistant
 */
class CalendarManager {
  constructor() {
    try {
      // Try multiple methods to get a calendar
      try {
        // First try the default calendar
        this.calendar = CalendarApp.getDefaultCalendar();
        console.log('Successfully initialized with default calendar:', this.calendar.getName());
      } catch (e) {
        console.log('Could not get default calendar, trying all calendars');
        // If default fails, try to get any calendar
        const calendars = CalendarApp.getAllCalendars();
        if (calendars && calendars.length > 0) {
          this.calendar = calendars[0];
          console.log('Using first available calendar:', this.calendar.getName());
        } else {
          throw new Error('No calendars available');
        }
      }
      
      // Initialize work hours from config
      this.workHours = getWorkHours();
      this.hasCalendarAccess = true;
    } catch (error) {
      console.error('Error initializing CalendarManager:', error);
      // Create a placeholder calendar object that won't break the code
      this.calendar = null;
      this.hasCalendarAccess = false;
    }
  }

  /**
   * Find available time slots within work hours
   * @param {number} duration - Duration in minutes
   * @param {Date} date - Date to check
   * @param {Object} workHours - Work hours configuration
   * @param {Array} existingEvents - Existing events to consider
   * @returns {Array<Object>} Available time slots
   */
  findAvailableSlots(duration, date, workHours, existingEvents) {
    // First validate parameters
    if (!(date instanceof Date)) {
      console.error('Invalid date parameter:', date);
      throw new Error('Second parameter must be a Date object');
    }
    if (typeof duration !== 'number' || duration <= 0) {
      console.error('Invalid duration:', duration);
      throw new Error('First parameter must be a positive number');
    }

    console.log(`Finding slots for ${duration} mins on ${date.toLocaleString()}`);
    
    // Create a new Date object to avoid modifying the input date
    const searchDate = new Date(date.getTime());
    
    // Convert day number to name
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const dayName = days[searchDate.getDay()];
    
    // Use passed in workHours if provided, otherwise use instance workHours
    const workHoursConfig = (workHours && workHours[dayName]) || this.workHours[dayName];
    
    console.log('Work hours config:', JSON.stringify(workHoursConfig));
    
    if (!workHoursConfig) {
      console.log(`No work hours for ${dayName}`);
      return [];
    }

    // Use passed in events if provided, otherwise get from calendar
    const events = existingEvents || this.calendar.getEvents(
      new Date(searchDate.getFullYear(), searchDate.getMonth(), searchDate.getDate(), 0, 0, 0),
      new Date(searchDate.getFullYear(), searchDate.getMonth(), searchDate.getDate(), 23, 59, 59)
    );
    
    console.log('Existing events:', events.length);

    // Filter events to only include those that should block time
    const blockingEvents = events.filter(event => {
      try {
        // Check if this is a custom tracking object
        if (this.isCustomEvent(event)) {
          return true; // Custom events are always blocking
        }
        
        // For Calendar events, check normally
        const isMyEvent = event.getCreators().includes('pkumar@eightfold.ai');
        const myStatus = event.getMyStatus();
        const isAccepted = myStatus === CalendarApp.GuestStatus.YES || 
                          myStatus === CalendarApp.GuestStatus.OWNER;
        
        return isMyEvent || isAccepted;
      } catch (e) {
        console.log('Error checking event:', e);
        return false;
      }
    });

    const slots = [];
    
    // Check each work hour range
    for (const range of workHoursConfig) {
      const { start: startStr, end: endStr } = range;
      const startTime = this.parseTime(startStr);
      const endTime = this.parseTime(endStr);
      
      // Start checking from the beginning of work hours
      let slotStart = new Date(searchDate);
      slotStart.setHours(startTime.hours, startTime.minutes, 0, 0);
      
      const rangeEnd = new Date(searchDate);
      rangeEnd.setHours(endTime.hours, endTime.minutes, 0, 0);

      // Check slots in 30-minute increments
      while (slotStart < rangeEnd) {
        const slotEnd = new Date(slotStart.getTime() + duration * 60000);
        
        // If slot would end after work hours, move to next range
        if (slotEnd > rangeEnd) {
          break;
        }

        // Skip slots that are in the past
        const now = new Date();
        if (slotStart < now) {
          // Move to next 30-minute slot
          slotStart = new Date(slotStart.getTime() + 30 * 60000);
          continue;
        }

        let hasOverlap = false;
        
        // Check for overlaps with blocking events
        for (const event of blockingEvents) {
          try {
            const eventStart = this.isCustomEvent(event) ? event.start : event.getStartTime();
            const eventEnd = this.isCustomEvent(event) ? event.end : event.getEndTime();
            const eventTitle = this.isCustomEvent(event) ? (event.title || event.name) : event.getTitle();
            
            // First check for exact time matches - prevent double booking
            if (slotStart.getTime() === eventStart.getTime()) {
              console.log(`Slot ${slotStart.toLocaleTimeString()} exact match with ${eventTitle}`);
              hasOverlap = true;
              break;
            }
            
            // Then check for any overlap
            if (!(slotEnd <= eventStart || slotStart >= eventEnd)) {
              console.log(`Slot ${slotStart.toLocaleTimeString()} - ${slotEnd.toLocaleTimeString()} overlaps with ${eventTitle}`);
              hasOverlap = true;
              break;
            }
          } catch (e) {
            console.log('Error checking event overlap:', e);
            continue;
          }
        }

        if (!hasOverlap) {
          slots.push({
            start: new Date(slotStart),
            end: new Date(slotEnd)
          });
        }

        // Move to next 30-minute slot
        slotStart = new Date(slotStart.getTime() + 30 * 60000);
      }
    }

    return slots;
  }

  /**
   * Clear all auto-scheduled events
   */
  clearAutoScheduledEvents() {
    const now = new Date();
    // Look back 1 day and forward 14 days to ensure we catch all auto-scheduled events
    const startDate = new Date(now);
    startDate.setDate(startDate.getDate() - 1);
    
    const endDate = new Date(now);
    endDate.setDate(endDate.getDate() + 14);
    
    console.log(`Searching for auto-scheduled events with title containing 'Auto-Scheduled:' from ${startDate.toLocaleString()} to ${endDate.toLocaleString()}`);
    
    const events = this.calendar.getEvents(startDate, endDate);
    console.log(`Found ${events.length} total events in search range`);
    
    let deletedCount = 0;
    events.forEach(event => {
      try {
        const title = event.getTitle();
        console.log(`Checking event: "${title}"`);
        if (title.indexOf('Auto-Scheduled:') !== -1) {
          console.log(`Deleting event: "${title}"`);
          event.deleteEvent();
          deletedCount++;
        }
      } catch (e) {
        console.log('Error processing event:', e);
      }
    });
    
    console.log(`Cleared ${deletedCount} previously auto-scheduled events`);
  }

  /**
   * Schedule a task on the calendar
   * @param {Object} task - Task to schedule
   * @param {Date} startTime - Start time
   * @param {Date} [endTime] - Optional end time (calculated from timeBlock if not provided)
   * @returns {Object} Scheduled event details
   */
  scheduleTask(task, startTime, endTime) {
    // Calculate end time if not provided
    const calculatedEndTime = endTime || new Date(startTime.getTime() + (task.timeBlock || 30) * 60000);
    
    // Create calendar event
    const event = this.calendar.createEvent(
      `Auto-Scheduled: ${task.name}`,
      startTime,
      calculatedEndTime,
      {
        description: `Priority: ${task.priority}\nNotes: ${task.notes || 'None'}`,
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
   * Get all events for a specific day
   * @param {Date} date - The date to check
   * @returns {Array} Array of events
   */
  getEventsForDay(date) {
    const calendar = CalendarApp.getDefaultCalendar();
    const startOfDay = new Date(date);
    startOfDay.setHours(0, 0, 0, 0);
    
    const endOfDay = new Date(date);
    endOfDay.setHours(23, 59, 59, 999);
    
    return calendar.getEvents(startOfDay, endOfDay);
  }

  /**
   * Parse time string (HH:MM) into hours and minutes
   * @param {string} timeStr - Time string in HH:MM format
   * @returns {Object} Hours and minutes
   */
  parseTime(timeStr) {
    const [hours, minutes] = timeStr.split(':').map(Number);
    return { hours, minutes };
  }

  /**
   * Check if an object is a custom event tracking object vs. a Google Calendar event
   * @param {Object} event - Event to check
   * @returns {boolean} True if custom object
   */
  isCustomEvent(event) {
    // A custom event has title, start and end properties
    // but doesn't have Calendar API methods
    return event && 
           typeof event.getCreators !== 'function' && 
           (event.title || event.name) && 
           event.start && 
           event.end;
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
    
    const now = new Date();
    const startDate = new Date(now);
    startDate.setDate(startDate.getDate() - 1);
    
    const endDate = new Date(now);
    endDate.setDate(endDate.getDate() + 14);
    
    console.log(`Searching for events with title "Auto-Scheduled: ${taskName}"`);
    
    const events = this.calendar.getEvents(startDate, endDate);
    let deletedCount = 0;
    
    events.forEach(event => {
      try {
        const title = event.getTitle();
        
        // Only delete events for this specific task
        if (title === `Auto-Scheduled: ${taskName}`) {
          console.log(`Deleting event: "${title}"`);
          event.deleteEvent();
          deletedCount++;
        }
      } catch (e) {
        console.log('Error processing event:', e);
      }
    });
    
    console.log(`Cleared ${deletedCount} events for task "${taskName}"`);
  }

  // Add a method to check if calendar is available
  isCalendarAvailable() {
    return this.calendar !== null;
  }

  /**
   * Create a calendar event for a task
   * @param {Object} task - Task to schedule
   * @param {Date} startTime - Start time for the event
   * @returns {CalendarEvent} Created event
   */
  createCalendarEvent(task, startTime) {
    try {
      // Validate task name
      const taskName = task.name || 'Unnamed Task';
      
      // Calculate end time based on time block
      const timeBlock = task.timeBlock || 30;
      const endTime = new Date(startTime.getTime() + timeBlock * 60 * 1000);
      
      // Create event
      const event = this.calendar.createEvent(
        `Auto-Scheduled: ${taskName}`,
        startTime,
        endTime,
        {
          description: `Priority: ${task.priority || 'Not set'}\nTime Block: ${timeBlock} minutes\n\nAuto-scheduled by Digital Assistant.`
        }
      );
      
      console.log(`Created calendar event "${taskName}" from ${startTime.toLocaleString()} to ${endTime.toLocaleString()}`);
      return event;
    } catch (error) {
      console.error('Error creating calendar event:', error);
      throw error;
    }
  }

  /**
   * Get existing calendar events for a given date range
   * @param {Date} [startDate] - Start date (defaults to now)
   * @param {Date} [endDate] - End date (defaults to 14 days from start)
   * @returns {Array} Array of calendar events
   */
  getExistingEvents(startDate = new Date(), endDate = new Date(startDate.getTime() + 14 * 24 * 60 * 60 * 1000)) {
    if (!this.isCalendarAvailable()) {
      console.warn('Calendar not available, returning empty array for getExistingEvents');
      return [];
    }

    try {
      // Format dates for logging
      const startStr = startDate.toLocaleString();
      const endStr = endDate.toLocaleString();
      console.log(`Getting events from ${startStr} to ${endStr}`);
      
      // Get events from the calendar
      const events = this.calendar.getEvents(startDate, endDate);
      
      // Log detailed information about each event
      console.log(`Retrieved ${events.length} existing events from calendar`);
      events.forEach(event => {
        const eventStart = event.getStartTime();
        const eventEnd = event.getEndTime();
        const eventTitle = event.getTitle();
        console.log(`Event: "${eventTitle}" from ${eventStart.toLocaleString()} to ${eventEnd.toLocaleString()}`);
      });
      
      return events;
    } catch (error) {
      console.error('Error getting existing events:', error);
      return [];
    }
  }
}

// Create global instance
const calendarManager = new CalendarManager();