/**
 * Script to fix Auto Update.js issues
 * This will replace only the problematic functions while keeping all other functionality intact
 * 
 * Note: Keep your existing Menu.js, Artist Breakdown.js, and Concert Analysis.js files - don't delete them!
 */

// Ensure we're using consistent sheet names
const AUTO_UPDATE_CONSTANTS = {
  SHEET_NAMES: {
    MAIN: "Tracker",
    CALENDAR: "Calendar View",
    MOBILE_CALENDAR: "Mobile Calendar",
    UPCOMING_MOBILE: "Upcoming Mobile Calendar"
  }
};

/**
 * This function gets called by the time-based trigger at midnight
 * Fixed version that properly updates all calendar views
 */
function runDailyUpdate() {
  try {
    Logger.log("Starting daily automatic update at " + new Date().toString());
    
    // Get a reference to the spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the main tracker sheet
    const mainSheet = ss.getSheetByName(AUTO_UPDATE_CONSTANTS.SHEET_NAMES.MAIN);
    if (!mainSheet) {
      Logger.log("Error: Main tracker sheet not found");
      return;
    }
    
    // Run the calendar functions directly
    try {
      // Use the existing calendar generation functions
      
      if (typeof generateDesktopCalendar === 'function') {
        generateDesktopCalendar();
        Logger.log("Desktop calendar updated successfully");
      } else {
        Logger.log("Error: generateDesktopCalendar function not found");
      }
      
      if (typeof generateMobileCalendar === 'function') {
        generateMobileCalendar();
        Logger.log("Mobile calendar updated successfully");
      } else {
        Logger.log("Error: generateMobileCalendar function not found");
      }
      
      if (typeof generateUpcomingMobileView === 'function') {
        generateUpcomingMobileView();
        Logger.log("Upcoming shows view updated successfully");
      } else {
        Logger.log("Error: generateUpcomingMobileView function not found");
      }
      
      // Also update the analysis dashboard if possible
      if (typeof updateAnalysis === 'function') {
        updateAnalysis();
        Logger.log("Analysis dashboard updated successfully");
      }

      // Update Artist Breakdown sheet
      if (typeof createArtistBreakdown === 'function') {
        createArtistBreakdown();
        Logger.log("Artist Breakdown updated successfully");
      }

      // Update Venue Breakdown sheet
      if (typeof createVenueBreakdown === 'function') {
        createVenueBreakdown();
        Logger.log("Venue Breakdown updated successfully");
      }
      
      // Add a note in a cell to show when the last update occurred
      const analysisSheet = ss.getSheetByName("Concert Analysis");
      if (analysisSheet) {
        const lastRow = analysisSheet.getLastRow();
        if (lastRow > 0) {
          const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy h:mm a");
          analysisSheet.getRange(lastRow + 2, 1).setValue("Last automatic update: " + timestamp);
        }
      }
    } catch (functionError) {
      Logger.log("Error in calendar function: " + functionError.toString());
    }
    
    // Log completion
    Logger.log("Daily update completed successfully at " + new Date().toString());
  } catch (error) {
    // Log any errors that might occur
    Logger.log("Error in daily update: " + error.toString());
  }
}

/**
 * Create a daily trigger to update all calendar views at 2:00 AM
 * Fixed version that ensures proper trigger setup
 */
function createDailyTrigger() {
  // Delete any existing daily update triggers to avoid duplicates
  deleteDailyUpdateTriggers();
  
  // Create a new daily trigger that runs at 2:00 AM
  ScriptApp.newTrigger('runDailyUpdate')
    .timeBased()
    .atHour(2) // Run at 2:00 AM (more likely to succeed than midnight)
    .everyDays(1) // Every day
    .create();
  
  // Safely confirm creation to user
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast("Daily update trigger created. Calendars will now update automatically at 2:00 AM.", "Auto-Update Enabled", 10);
  } catch (e) {
    // In case we're not in a UI context
    Logger.log("Daily update trigger created, but couldn't display toast notification");
  }
  
  // Also save the auto-update state to document properties
  PropertiesService.getDocumentProperties().setProperty('autoUpdateEnabled', 'true');
  PropertiesService.getDocumentProperties().setProperty('autoUpdateSetOn', new Date().toString());
  
  // Log the setup for troubleshooting
  Logger.log("Daily update trigger created: Calendars will update at 2:00 AM daily");
}

/**
 * Delete any existing daily update triggers
 */
function deleteDailyUpdateTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  
  for (let i = 0; i < allTriggers.length; i++) {
    const trigger = allTriggers[i];
    // Check if it's our daily update trigger
    if (trigger.getHandlerFunction() === 'runDailyUpdate') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("Deleted existing daily update trigger");
    }
  }
}

/**
 * Disable the daily update by removing all associated triggers
 */
function disableDailyUpdate() {
  deleteDailyUpdateTriggers();
  
  // Safely confirm to user
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast("Daily updates have been disabled.", "Auto-Update Disabled", 5);
  } catch (e) {
    // In case we're not in a UI context
    Logger.log("Daily updates disabled, but couldn't display toast notification");
  }
  
  // Update the auto-update state in document properties
  PropertiesService.getDocumentProperties().setProperty('autoUpdateEnabled', 'false');
  PropertiesService.getDocumentProperties().setProperty('autoUpdateDisabledOn', new Date().toString());
  
  Logger.log("Daily updates disabled");
}

/**
 * Check if daily update is enabled
 * @return {boolean} True if daily update is enabled
 */
function isDailyUpdateEnabled() {
  const allTriggers = ScriptApp.getProjectTriggers();
  
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'runDailyUpdate') {
      return true;
    }
  }
  
  return false;
}

/**
 * Setup function to replace your current Auto Update.js content
 * DO NOT delete your other files - just add this one
 */
function setupAutoUpdate() {
  const ui = SpreadsheetApp.getUi();
  
  // Confirm with user
  const response = ui.alert(
    'Setup Auto Update',
    'This will set up a trigger to automatically update your calendar views at 2:00 AM daily.\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Setup Cancelled', 'No changes were made.', ui.ButtonSet.OK);
    return;
  }
  
  // Set up the trigger
  createDailyTrigger();
  
  // Show success message
  ui.alert(
    'Setup Successful',
    'Your concert calendars will now update automatically at 2:00 AM daily.\n\n' +
    'You can still update them manually from the "Calendar Views" menu.',
    ui.ButtonSet.OK
  );
}
