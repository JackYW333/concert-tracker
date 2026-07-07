/**
 * Script to automatically update analysis dashboard and artist breakdown
 * Runs automatically when a new "Attended" entry is added or at 2:00 AM daily
 */

// Option 1: Run whenever a new "Attended" entry is added
// Modify the onEdit function to include Venue Breakdown
function onEdit(e) {
  // Only proceed if an edit was made to the main tracker sheet
  if (!e || !e.range || e.range.getSheet().getName() !== "Tracker") {
    return;
  }
  
  // Check if the edit was in the Status column and is set to "Attended"
  const sheet = e.range.getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf("Status");
  
  // If not editing the Status column, exit
  if (statusCol === -1 || e.range.getColumn() !== statusCol + 1) {
    return;
  }
  
  // Check if the new value is "Attended"
  if (e.value === "Attended") {
    try {
      const row = e.range.getRow();
      const artistCol = headers.indexOf("Artist");
      const venueCol = headers.indexOf("Venue");

      // Auto-fetch setlist URL before analysis so the artist breakdown picks it up
      autoFetchSetlistUrl(sheet, row, headers);

      // Update Analysis Dashboard
      updateAnalysis();

      // Update Artist Breakdown
      if (artistCol !== -1) {
        const artistName = sheet.getRange(row, artistCol + 1).getValue();

        // Create artist breakdown if needed
        createArtistBreakdown();

        // If an artist was edited, analyze that artist specifically
        if (artistName) {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          const artistSheet = ss.getSheetByName("Artist Breakdown");

          if (artistSheet) {
            artistSheet.getRange("B5").setValue(artistName);
            analyzeSelectedArtist();
          }
        }
      }

      // Update Venue Breakdown
      if (venueCol !== -1) {
        const venueName = sheet.getRange(row, venueCol + 1).getValue();

        // Create venue breakdown if needed
        createVenueBreakdown();

        if (venueName) {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          const venueSheet = ss.getSheetByName("Venue Breakdown");

          if (venueSheet) {
            venueSheet.getRange("B5").setValue(venueName);
            analyzeSelectedVenue();
          }
        }
      }

      // Bust the server-side cache so the next web app request reflects the new data
      incrementCacheGeneration();

      SpreadsheetApp.getActiveSpreadsheet().toast("Analysis dashboards updated automatically", "Auto Update");
    } catch (error) {
      console.error("Error in automatic update:", error);
    }
  }
}

// Function to set up all needed triggers
function setupAllTriggers() {
  // Set up edit trigger
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clean up existing triggers first
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'onEdit' || 
        allTriggers[i].getHandlerFunction() === 'runDailyUpdate') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
  
  // Create onEdit trigger for immediate updates
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  
  // Create daily trigger at 2:00 AM
  ScriptApp.newTrigger('runDailyUpdate')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();
  
  SpreadsheetApp.getActiveSpreadsheet().toast("All automatic update triggers have been set up", "Success");
}