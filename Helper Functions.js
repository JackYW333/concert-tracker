/**
 * Helper functions to support automatic triggers
 */

/**
 * Check if a specific Artist Breakdown exists and create it if not
 * This ensures the artist breakdown functionality is available for auto-triggers
 */
function ensureArtistBreakdownExists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const artistSheet = ss.getSheetByName("Artist Breakdown");
  
  // If sheet doesn't exist, create it
  if (!artistSheet) {
    createArtistBreakdown();
    return true;
  }
  
  return true;
}

/**
 * Check for new "Attended" entries since last update by comparing counts.
 * Useful for daily triggers to only update when new shows were added.
 */
function hasNewAttendedEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  if (!mainSheet) return false;

  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf("Status");
  if (statusCol === -1) return false;

  const currentCount = data.slice(1).filter(row => row[statusCol] === "Attended").length;
  const props = PropertiesService.getScriptProperties();
  const lastCount = parseInt(props.getProperty("lastAttendedCount") || "0");

  if (currentCount > lastCount) {
    props.setProperty("lastAttendedCount", currentCount.toString());
    return true;
  }
  return false;
}

/**
 * Sort an array of "Mon YYYY" month strings chronologically.
 * @param {string[]} months - Array of month strings (e.g. ["Jan 2025", "Mar 2024"])
 * @return {string[]} Sorted array
 */
function sortMonthsChronologically(months) {
  return months.sort((a, b) => {
    const [aMonth, aYear] = a.split(' ');
    const [bMonth, bYear] = b.split(' ');
    const yearDiff = parseInt(aYear) - parseInt(bYear);
    if (yearDiff !== 0) return yearDiff;
    return MONTHS.indexOf(aMonth) - MONTHS.indexOf(bMonth);
  });
}

/**
 * Optimized version of analyzeSelectedArtist that can be called by triggers
 * @param {string} artistName - Name of the artist to analyze (optional)
 */
function analyzeArtist(artistName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Make sure Artist Breakdown sheet exists
  ensureArtistBreakdownExists();
  
  const artistSheet = ss.getSheetByName("Artist Breakdown");
  
  if (!artistName) {
    // Try to get artist from the dropdown
    artistName = artistSheet.getRange("B5").getValue();
  }
  
  if (!artistName) {
    // No artist selected
    return false;
  }
  
  // Set the artist in the dropdown
  artistSheet.getRange("B5").setValue(artistName);
  
  // Run analysis for this artist
  analyzeSelectedArtist();
  
  return true;
}

/**
 * Helper function to analyze a venue programmatically
 * Useful for automatic triggers and batch operations
 * @param {string} venueName - Name of the venue to analyze (optional)
 * @returns {boolean} - True if analysis was performed, false otherwise
 */
function analyzeVenue(venueName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Make sure Venue Breakdown sheet exists
  const venueSheet = ss.getSheetByName("Venue Breakdown");
  
  if (!venueSheet) {
    // Create it if it doesn't exist
    createVenueBreakdown();
  }
  
  // Get the venue sheet again (in case it was just created)
  const updatedVenueSheet = ss.getSheetByName("Venue Breakdown");
  
  if (!venueName) {
    // Try to get venue from the dropdown
    venueName = updatedVenueSheet.getRange("B5").getValue();
  }
  
  if (!venueName) {
    // No venue selected
    return false;
  }
  
  // Set the venue in the dropdown
  updatedVenueSheet.getRange("B5").setValue(venueName);
  
  // Run analysis for this venue
  analyzeSelectedVenue();
  
  return true;
}