/**
 * Sheet operations for Digital Assistant
 */
class SheetManager {
  constructor() {
    this.ss = SpreadsheetApp.getActive();
    this.tasksSheet = null; // Will be initialized when needed
  }

  /**
   * Get or create Tasks sheet
   * @returns {Sheet} Tasks sheet
   */
  getTasksSheet() {
    if (!this.tasksSheet) {
      this.tasksSheet = this.ss.getSheetByName('Tasks');
      if (!this.tasksSheet) {
        this.tasksSheet = this.ss.insertSheet('Tasks');
      }
    }
    return this.tasksSheet;
  }

  /**
   * Initialize all required sheets
   */
  initializeSheets() {
    this.setupTasksSheet();
  }

  /**
   * Setup Tasks sheet with required columns
   */
  setupTasksSheet() {
    const tasksSheet = this.getTasksSheet();
    
    // Setup columns
    const headers = [
      ['Tasks', 'Priority', 'Time Block', 'Deadline', 'Status', 'Notes']
    ];
    
    // Clear existing content
    tasksSheet.clear();
    
    // Write headers
    tasksSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    
    // Format headers
    tasksSheet.getRange(1, 1, 1, headers[0].length)
      .setFontWeight('bold')
      .setBackground('#D3D3D3')
      .setHorizontalAlignment('center');
    
    // Set column widths
    tasksSheet.setColumnWidths(1, headers[0].length, 150);
    
    // Add data validation for Priority - UPDATED to include Follow-up
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['P1', 'P2', 'P3', 'Follow-up'], true)
      .setAllowInvalid(false)
      .build();
    tasksSheet.getRange('B2:B1000').setDataValidation(priorityRule);
    
    // Add data validation for Status
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Scheduled', 'Done', 'Pause'], true)
      .setAllowInvalid(false)
      .build();
    tasksSheet.getRange('E2:E1000').setDataValidation(statusRule);
  }

  /**
   * Add a new task to the sheet
   * @param {string|Object} taskNameOrObject - Name of the task or task object
   * @param {string} [priority] - Priority of the task
   * @param {number} [timeBlock] - Time block in minutes
   * @param {string} [notes] - Additional notes
   * @returns {Object} Result with row number
   */
  addTask(taskNameOrObject, priority, timeBlock, notes) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
      
      // Handle both object and individual parameters
      let taskName, taskPriority, taskTimeBlock, taskNotes;
      
      if (typeof taskNameOrObject === 'object' && taskNameOrObject !== null) {
        // If first parameter is an object, extract properties
        const taskObj = taskNameOrObject;
        taskName = taskObj.name || '';
        taskPriority = taskObj.priority || 'P2';
        taskTimeBlock = taskObj.timeBlock || 30;
        taskNotes = taskObj.notes || '';
        
        console.log(`Processing task object: Name="${taskName}", Priority="${taskPriority}"`);
      } else {
        // If individual parameters, use them directly
        taskName = taskNameOrObject || '';
        taskPriority = priority || 'P2';
        taskTimeBlock = timeBlock || 30;
        taskNotes = notes || '';
      }
      
      // ENHANCED LOGGING: Log call stack to identify where this is being called from
      const stack = new Error().stack;
      console.log(`TASK CREATION TRACE - Adding task "${taskName}" - Call stack: ${stack}`);
      
      // Get headers to find column positions
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // Find column indices
      const taskColIndex = headers.findIndex(header => 
        String(header).toLowerCase() === 'tasks'
      );
      const priorityColIndex = headers.findIndex(header => 
        String(header).toLowerCase() === 'priority'
      );
      const timeBlockColIndex = headers.findIndex(header => 
        String(header).toLowerCase() === 'time block'
      );
      const notesColIndex = headers.findIndex(header => 
        String(header).toLowerCase() === 'notes'
      );
      const statusColIndex = headers.findIndex(header => 
        String(header).toLowerCase() === 'status'
      );
      
      // Validate column indices
      if (taskColIndex === -1) throw new Error('Tasks column not found');
      
      // DUPLICATE CHECK: Check if this task already exists
      const allData = sheet.getDataRange().getValues();
      let existingTaskRow = -1;
      
      // Skip header row (index 0)
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][taskColIndex] === taskName) {
          existingTaskRow = i + 1; // +1 because spreadsheet rows are 1-indexed
          console.log(`DUPLICATE DETECTION: Task "${taskName}" already exists at row ${existingTaskRow}`);
          break;
        }
      }
      
      // If task already exists, use existing row instead of creating a duplicate
      if (existingTaskRow > 0) {
        console.log(`DUPLICATE HANDLING: Task "${taskName}" exists at row ${existingTaskRow}. Returning existing row.`);
        
        // Check the status of the existing task
        const existingStatus = (statusColIndex !== -1) ? allData[existingTaskRow-1][statusColIndex] : '';
        console.log(`Existing task status: "${existingStatus}"`);
        
        // Return the existing row information
        return {
          success: true,
          row: existingTaskRow,
          message: "Using existing task",
          isExisting: true
        };
      }
      
      // Get the next available row
      const lastRow = sheet.getLastRow();
      const newRow = lastRow + 1;
      
      console.log(`SPREADSHEET STATE: Last row before adding: ${lastRow}, New row: ${newRow}`);
      
      // Prepare row data
      const rowData = [];
      for (let i = 0; i < headers.length; i++) {
        if (i === taskColIndex) {
          rowData.push(taskName);
        } else if (i === priorityColIndex && priorityColIndex !== -1) {
          rowData.push(taskPriority);
        } else if (i === timeBlockColIndex && timeBlockColIndex !== -1) {
          rowData.push(taskTimeBlock);
        } else if (i === notesColIndex && notesColIndex !== -1) {
          rowData.push(taskNotes);
        } else if (i === statusColIndex && statusColIndex !== -1) {
          rowData.push('Pending');
        } else {
          rowData.push(''); // Empty for other columns
        }
      }
      
      // Add the row
      sheet.appendRow(rowData);
      
      // Verify the row was actually added
      const newLastRow = sheet.getLastRow();
      console.log(`VERIFICATION: After adding, sheet has ${newLastRow} rows (expected ${newRow})`);
      
      console.log(`Added task "${taskName}" with priority "${taskPriority}" at row ${newRow}`);
      
      return {
        success: true,
        row: newRow,
        isExisting: false
      };
    } catch (error) {
      console.error('Error adding task:', error);
      throw error;
    }
  }

  /**
   * Update task status
   * @param {number} row - Row number
   * @param {string} status - New status
   * @returns {boolean} Success
   */
  updateTaskStatus(row, status) {
    try {
      console.log(`Updating task status at row ${row} to "${status}"`);
      
      const sheet = this.getTasksSheet();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // Find the status column
      const statusColIndex = headers.findIndex(header => 
        String(header).toLowerCase() === 'status'
      );
      
      if (statusColIndex < 0) {
        console.error('Status column not found');
        return false;
      }
      
      // Update the status
      sheet.getRange(row, statusColIndex + 1).setValue(status);
      console.log(`Successfully updated status at row ${row}, column ${statusColIndex + 1} to "${status}"`);
      
      return true;
    } catch (error) {
      console.error(`Error updating task status at row ${row}:`, error);
      return false;
    }
  }

  /**
   * Get all tasks from the sheet
   * @returns {Array<Object>} Array of task objects
   */
  getTasks() {
    const tasksSheet = this.getTasksSheet();
    const lastRow = tasksSheet.getLastRow();
    if (lastRow <= 1) return []; // Only headers exist
    
    const data = tasksSheet.getRange(2, 1, lastRow - 1, tasksSheet.getLastColumn()).getValues(); // Read all columns
    const headers = tasksSheet.getRange(1, 1, 1, tasksSheet.getLastColumn()).getValues()[0];

    // Find column indices based on headers
    const nameCol = headers.findIndex(h => h.toLowerCase() === 'tasks');
    const priorityCol = headers.findIndex(h => h.toLowerCase() === 'priority');
    const timeBlockCol = headers.findIndex(h => h.toLowerCase() === 'time block');
    const deadlineCol = headers.findIndex(h => h.toLowerCase() === 'deadline');
    const statusCol = headers.findIndex(h => h.toLowerCase() === 'status');
    const notesCol = headers.findIndex(h => h.toLowerCase() === 'notes');
    // Add scheduledTimeCol if needed, or assume it might not exist
    const scheduledTimeCol = headers.findIndex(h => h.toLowerCase() === 'scheduled time'); 

    return data.map((row, index) => {
      // Read priority first to use in status defaulting
      const priorityValue = priorityCol !== -1 ? (row[priorityCol] || 'P1') : 'P1';
      const priorityLower = String(priorityValue).toLowerCase().trim();

      // Read status value from sheet
      const statusValue = statusCol !== -1 ? row[statusCol] : '';
      let finalStatus;

      // Determine final status based on priority and sheet value
      if (!statusValue) { // If status cell is blank
        if (priorityLower === 'follow-up' || priorityLower === 'follow up') {
          finalStatus = 'Follow-up'; // Default Status to Follow-up if Priority is Follow-up
        } else {
          finalStatus = 'Pending'; // Otherwise default blank status to Pending
        }
      } else {
        finalStatus = statusValue; // Use the value from the sheet if not blank
      }

      return {
        name: nameCol !== -1 ? row[nameCol] : '',
        priority: priorityValue, // Use the already determined priorityValue
        timeBlock: timeBlockCol !== -1 ? (row[timeBlockCol] || 30) : 30, // Default 30
        deadline: deadlineCol !== -1 ? row[deadlineCol] : '',
        status: finalStatus, // Use the determined finalStatus
        notes: notesCol !== -1 ? row[notesCol] : '',
        scheduledTime: scheduledTimeCol !== -1 ? row[scheduledTimeCol] : null, // Handle missing column
        row: index + 2 // Adding 2 because: 1 for header, 1 for zero-based index
      };
    });
  }

  /**
   * Get tasks by status
   * @param {string|Array} status - Status to filter by (can be string or array of strings)
   * @returns {Array} Filtered tasks
   */
  getTasksByStatus(status) {
    const tasks = this.getTasks();
    const statusArray = Array.isArray(status) ? status : [status];
    return tasks.filter(task => statusArray.includes(task.status || ''));
  }

  /**
   * Get tasks filtered by priority
   * @param {string} priority - Priority to filter by ('P1', 'P2', 'P3')
   * @returns {Array<Object>} Filtered tasks
   */
  getTasksByPriority(priority) {
    const allTasks = this.getTasks();
    return allTasks.filter(task => task.priority === priority);
  }

  /**
   * Get a task by its row number
   * @param {number} row - Row number
   * @returns {Object} Task object
   */
  getTaskByRow(row) {
    try {
      const sheet = this.getTasksSheet();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // Create task object with all properties
      const task = {
        row: row,
        name: '',
        priority: '',
        timeBlock: 30,
        deadline: '',
        status: 'Pending',
        notes: ''
      };
      
      // Map data from row to task object
      headers.forEach((header, index) => {
        const headerLower = String(header).toLowerCase();
        const value = rowData[index];
        
        if (headerLower === 'tasks' || headerLower === 'task') {
          task.name = value;
        } else if (headerLower === 'priority') {
          task.priority = value;
        } else if (headerLower === 'time block') {
          task.timeBlock = value || 30;
        } else if (headerLower === 'deadline') {
          task.deadline = value;
        } else if (headerLower === 'status') {
          task.status = value || 'Pending';
        } else if (headerLower === 'notes') {
          task.notes = value;
        }
      });
      
      // Log the task for debugging
      console.log(`Retrieved task from row ${row}:`, JSON.stringify(task));
      
      return task;
    } catch (error) {
      console.error(`Error getting task from row ${row}:`, error);
      return null;
    }
  }

  /**
   * Remove duplicate tasks with the same name
   * @returns {Object} Result with count of removed duplicates
   */
  removeDuplicateTasks() {
    try {
      const sheet = this.getTasksSheet();
      
      // Get all data including headers
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      // Find column indices
      const taskColIndex = headers.findIndex(header => 
        String(header).toLowerCase() === 'tasks'
      );
      
      if (taskColIndex === -1) {
        console.error('Tasks column not found');
        return { success: false, message: 'Tasks column not found', count: 0 };
      }
      
      // Skip header row
      const taskRows = data.slice(1);
      
      // Map of task names to their newest row index
      const taskMap = new Map();
      
      // First pass: find newest row for each task name
      taskRows.forEach((row, index) => {
        const taskName = row[taskColIndex];
        if (taskName) {
          // Store the data row index + 2 (to account for header row and 1-indexing)
          taskMap.set(taskName, Math.max(index + 2, taskMap.get(taskName) || 0));
        }
      });
      
      // Find duplicate rows to delete
      const rowsToDelete = [];
      
      // Second pass: identify rows to delete
      taskRows.forEach((row, index) => {
        const taskName = row[taskColIndex];
        if (taskName) {
          const newestRow = taskMap.get(taskName);
          // If this is not the newest row for this task, mark for deletion
          if (index + 2 !== newestRow) {
            rowsToDelete.push(index + 2);
          }
        }
      });
      
      // Sort in descending order to avoid shifting issues when deleting
      rowsToDelete.sort((a, b) => b - a);
      
      // Delete rows
      for (const row of rowsToDelete) {
        console.log(`Deleting duplicate task row: ${row}`);
        sheet.deleteRow(row);
      }
      
      console.log(`Removed ${rowsToDelete.length} duplicate task(s)`);
      
      return {
        success: true,
        message: `Removed ${rowsToDelete.length} duplicate task(s)`,
        count: rowsToDelete.length
      };
    } catch (error) {
      console.error('Error removing duplicate tasks:', error);
      return {
        success: false,
        message: error.message,
        count: 0
      };
    }
  }
}

// Create global instance
const sheetManager = new SheetManager();

/**
 * Handle new task additions via installable trigger
 * @param {Event} e - The onEdit event
 */
// REMOVED: This function has been moved to Config.gs with improved functionality
// function onTaskEditTrigger(e) {
//   const sheet = e.source.getActiveSheet();
//   if (sheet.getName() !== 'Tasks') return;
//   
//   const row = e.range.getRow();
//   if (row === 1) return; // Skip header row
//   
//   const column = e.range.getColumn();
//   const newValue = e.value;
//   
//   // If new task added (column 1) or status changed to "Pending" (column 5)
//   if (column === 1 || (column === 5 && newValue === 'Pending')) {
//     // Process task immediately instead of using time-based trigger
//     try {
//       const task = sheetManager.getTasks().find(t => t.row === row);
//       if (task) {
//         taskManager.initialize(); // Ensure initialized
//         taskManager.processPendingTasks();
//       }
//     } catch (error) {
//       console.error('Error processing task:', error);
//     }
//   }
// } 

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
 * Update priority data validation to include Follow-up
 */
function updatePriorityValidation() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
    if (!sheet) {
      ui.alert('Error', 'Tasks sheet not found.', ui.ButtonSet.OK);
      return false;
    }
    
    // Find the priority column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const priorityColIndex = headers.findIndex(header => 
      String(header).toLowerCase() === 'priority'
    );
    
    if (priorityColIndex < 0) {
      ui.alert('Error', 'Priority column not found.', ui.ButtonSet.OK);
      return false;
    }
    
    // Create updated validation rule
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['P1', 'P2', 'P3', 'Follow-up'], true)
      .setAllowInvalid(false)
      .build();
    
    // Apply to the priority column
    const column = priorityColIndex + 1;
    sheet.getRange(2, column, sheet.getMaxRows() - 1, 1).setDataValidation(priorityRule);
    
    ui.alert('Success', 'Priority validation updated to include Follow-up tasks.', ui.ButtonSet.OK);
    return true;
  } catch (error) {
    console.error('Error updating priority validation:', error);
    ui.alert('Error', 'Failed to update priority validation: ' + error.message, ui.ButtonSet.OK);
    return false;
  }
} 

/**
 * Fix existing follow-up tasks with incorrect capitalization
 */
function standardizeFollowUpPriorities() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
    if (!sheet) {
      ui.alert('Error', 'Tasks sheet not found.', ui.ButtonSet.OK);
      return false;
    }
    
    // Get all data
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find priority column index
    const priorityColIndex = headers.findIndex(header => 
      String(header).toLowerCase() === 'priority'
    );
    
    if (priorityColIndex < 0) {
      ui.alert('Error', 'Priority column not found.', ui.ButtonSet.OK);
      return false;
    }
    
    // Count of fixed tasks
    let fixedCount = 0;
    
    // Check each row for follow-up tasks with incorrect capitalization
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const priority = row[priorityColIndex];
      
      // If priority is a follow-up variant but not exactly "Follow-up"
      if (priority && String(priority).toLowerCase().includes('follow')) {
        if (priority !== 'Follow-up') {
          // Set priority to the standardized "Follow-up"
          sheet.getRange(i + 1, priorityColIndex + 1).setValue('Follow-up');
          fixedCount++;
        }
      }
    }
    
    if (fixedCount > 0) {
      ui.alert('Success', `Standardized ${fixedCount} follow-up tasks to use consistent "Follow-up" format.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Info', 'No follow-up tasks needed standardization.', ui.ButtonSet.OK);
    }
    
    return true;
  } catch (error) {
    console.error('Error standardizing follow-up tasks:', error);
    ui.alert('Error', 'Failed to standardize follow-up tasks: ' + error.message, ui.ButtonSet.OK);
    return false;
  }
} 