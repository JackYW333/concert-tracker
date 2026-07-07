/**
 * Enhanced Artist breakdown functionality for Concert Tracker
 * Improved with a blue color scheme and better visual formatting
 * Now with setlist integration features and event name/headliner details
 */

// Constants specific to this module
const ARTIST_SHEET_NAME = "Artist Breakdown";

// Color scheme constants
const COLORS = {
  PRIMARY: "#1E88E5",       // Primary blue
  PRIMARY_LIGHT: "#BBDEFB", // Light blue for backgrounds
  PRIMARY_DARK: "#0D47A1",  // Dark blue for headers
  SECONDARY: "#E3F2FD",     // Very light blue for alternating rows
  TEXT_PRIMARY: "#212121",  // Almost black
  TEXT_SECONDARY: "#757575", // Medium gray
  ACCENT: "#2962FF",        // Accent blue for highlights
  DIVIDER: "#BDBDBD"        // Light gray for dividers
  // Rating colors now defined in getRatingColor() function for more gradual scale
};

/**
 * Set a gradient color for a rating value - more gradual scale
 * @param {number} rating - The rating value (1-10)
 * @return {string} - Hex color code
 */
function getRatingColor(rating) {
  if (!rating || isNaN(rating)) return null;
  
  rating = parseFloat(rating);
  
  // Create a more gradual color scale (using HSL color scheme under the hood)
  if (rating <= 1) return "#F44336"; // Red
  if (rating <= 2) return "#F4655C"; // Red-orange
  if (rating <= 3) return "#F48771"; // Orange-red
  if (rating <= 4) return "#F4A987"; // Light orange
  if (rating <= 5) return "#F4CA9E"; // Amber
  if (rating <= 6) return "#E6D576"; // Yellow-green
  if (rating <= 7) return "#B5D86A"; // Light green
  if (rating <= 8) return "#7EC850"; // Green
  if (rating <= 9) return "#5CB8C8"; // Teal
  return "#2196F3";                  // Blue
}

/**
 * Creates the Artist Breakdown sheet and analyzes artist data
 */
function createArtistBreakdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get main tracker sheet (using your existing sheet name)
  const mainSheet = ss.getSheetByName("Tracker");
  if (!mainSheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: Tracker sheet not found", "Error");
    return;
  }
  
  // Create or get the Artist Breakdown sheet
  let artistSheet = ss.getSheetByName(ARTIST_SHEET_NAME);
  if (!artistSheet) {
    artistSheet = ss.insertSheet(ARTIST_SHEET_NAME);
  }
  
  // Clear the sheet
  artistSheet.clear();
  
  // Get all concert data
  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indexes
  const artistCol = headers.indexOf("Artist");
  const statusCol = headers.indexOf("Status");
  
  // Filter for "Attended" shows only
  const attendedShows = data.slice(1).filter(row => row[statusCol] === "Attended");
  
  // Extract unique artists
  const artists = [...new Set(attendedShows.map(row => row[artistCol]))].sort();
  
  // Set up sheet header with styling - more compact for smaller screens
  const headerRange = artistSheet.getRange("A1:G1");
  headerRange.merge()
    .setValue("ARTIST BREAKDOWN DASHBOARD")
    .setFontWeight("bold")
    .setFontSize(16) // Slightly smaller font
    .setBackground(COLORS.PRIMARY)
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  artistSheet.setRowHeight(1, 40); // Reduced height
  
  // Add a subtitle/description
  const subtitleRange = artistSheet.getRange("A2:G2");
  subtitleRange.merge()
    .setValue("Analyse your concert history by artist")
    .setFontStyle("italic")
    .setBackground(COLORS.PRIMARY_LIGHT)
    .setHorizontalAlignment("center");
  
  // Create a visually appealing control section - more compact
  const controlBgRange = artistSheet.getRange("A4:G6");
  controlBgRange.setBackground(COLORS.SECONDARY);
  controlBgRange.setBorder(true, true, true, true, null, null, COLORS.DIVIDER, SpreadsheetApp.BorderStyle.SOLID);
  
  // Add select dropdown label with styling
  artistSheet.getRange("A5")
    .setValue("Select an artist:")
    .setFontWeight("bold")
    .setFontColor(COLORS.TEXT_PRIMARY)
    .setHorizontalAlignment("right")
    .setVerticalAlignment("middle");
  
  // Create dropdown list of artists
  const dropdown = artistSheet.getRange("B5");
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(artists).build();
  dropdown.setDataValidation(rule);
  dropdown.setBorder(true, true, true, true, null, null, COLORS.PRIMARY, SpreadsheetApp.BorderStyle.SOLID);
  
  // Add a styled analyse button
  const buttonRange = artistSheet.getRange("C5");
  buttonRange.setValue("Analyse")
    .setBackground(COLORS.PRIMARY) 
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setBorder(true, true, true, true, null, null, COLORS.PRIMARY_DARK, SpreadsheetApp.BorderStyle.SOLID);
  
  // Add instructions text - more compact
  artistSheet.getRange("D5").setValue("← Select & analyse")
    .setFontStyle("italic")
    .setFontColor(COLORS.TEXT_SECONDARY)
    .setVerticalAlignment("middle");
  
  // Set column widths - more compact for smaller screens
  artistSheet.setColumnWidth(1, 80);    // First column (labels) - even shorter
  artistSheet.setColumnWidth(2, 120);   // Second column (values or names)
  artistSheet.setColumnWidth(3, 80);    // Third column
  artistSheet.setColumnWidth(4, 80);    // Fourth column
  artistSheet.setColumnWidth(5, 120);   // Fifth column (companions)
  artistSheet.setColumnWidth(6, 120);   // Sixth column (for notes)
  
  // Add a divider
  const dividerRange = artistSheet.getRange("A7:E7");
  dividerRange.setBackground(COLORS.DIVIDER);
  artistSheet.setRowHeight(7, 2);
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Select an artist from the dropdown and click 'Analyse'", "Artist Breakdown Ready");
}

/**
 * Creates a formatted section header
 * @param {Sheet} sheet - The sheet to add header to
 * @param {number} row - The row to add header
 * @param {string} title - The header text
 * @return {number} - The next row after the header
 */
function createSectionHeader(sheet, row, title) {
  // More compact version - use full width for headers
  const headerRange = sheet.getRange(row, 1, 1, 7);
  headerRange.merge()
    .setValue(title)
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(COLORS.PRIMARY_DARK)
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 25); // Slightly reduced height
  
  return row + 1;
}

/**
 * Creates a table header row with specified columns
 * @param {Sheet} sheet - The sheet to add header to
 * @param {number} row - The row to add header
 * @param {Array} columns - Array of column headers
 * @return {number} - The next row after the header
 */
function createTableHeader(sheet, row, columns) {
  const headerRange = sheet.getRange(row, 1, 1, columns.length);
  headerRange.setValues([columns])
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY)
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  return row + 1;
}

/**
 * Analyze the selected artist when the button is clicked
 */
function analyzeSelectedArtist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const artistSheet = ss.getSheetByName(ARTIST_SHEET_NAME);
  
  if (!artistSheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Artist Breakdown sheet not found", "Error");
    return;
  }
  
  // Get the selected artist
  const selectedArtist = artistSheet.getRange("B5").getValue();
  
  if (!selectedArtist) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Please select an artist first", "No Artist Selected");
    return;
  }
  
  // Clear any previous analysis (below row 7)
  const lastRow = artistSheet.getLastRow();
  if (lastRow > 7) {
    artistSheet.getRange(8, 1, lastRow - 7, 10).clear();
  }
  
  // Show loading indicator - more compact
  const loadingRange = artistSheet.getRange("A8:G8");
  loadingRange.merge()
    .setValue(`Analysing ${selectedArtist}...`)
    .setFontStyle("italic")
    .setBackground(COLORS.SECONDARY);
  
  // Run the analysis
  populateArtistAnalysis(selectedArtist);
}

/**
 * Populate the analysis for the selected artist
 */
function populateArtistAnalysis(artistName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  const artistSheet = ss.getSheetByName(ARTIST_SHEET_NAME);
  
  if (!mainSheet || !artistSheet) {
    return;
  }
  
  // Get all concert data
  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indexes
  const artistCol = headers.indexOf("Artist");
  const dateCol = headers.indexOf("Date");
  const venueCol = headers.indexOf("Venue");
  const cityCol = headers.indexOf("City");
  const roleCol = headers.indexOf("Role");
  const eventNameCol = headers.indexOf("Event Name");
  const ratingCol = headers.indexOf("Rating");
  const notesCol = headers.indexOf("Notes");
  const withCol = headers.indexOf("Who I Went With");
  const statusCol = headers.indexOf("Status");
  
  // Look for Setlist column either by name or by position (Column N is index 13)
  let setlistCol = headers.indexOf("Setlist");
  if (setlistCol === -1) {
    // Try index 13 (Column N) as fallback
    setlistCol = 13;
    Logger.log("Setlist column not found by name, using index 13 (Column N) instead");
  }
  
  // Filter for shows with this artist
  const artistShows = data.slice(1).filter(row => 
    row[artistCol] === artistName && row[statusCol] === "Attended"
  );
  
  // If no shows found, display message - more compact
  if (artistShows.length === 0) {
    artistSheet.getRange("A8:G9").merge()
      .setValue(`No attended shows found for ${artistName}`)
      .setBackground("#FFF9C4")  // Light yellow warning background
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
    return;
  }
  
  // Clear any loading indicator
  artistSheet.getRange("A8:G8").clear();
  
  // Set title with artist name - more compact
  let row = 8;
  const titleRange = artistSheet.getRange(row, 1, 1, 7);
  titleRange.merge()
    .setValue(`Analysis for ${artistName}`)
    .setFontWeight("bold")
    .setFontSize(14) // Reduced font size
    .setBackground(COLORS.PRIMARY)
    .setFontColor("white")
    .setHorizontalAlignment("center");
  row += 2; // Space after title
  
  // SUMMARY STATISTICS
  row = createSectionHeader(artistSheet, row, "SUMMARY STATISTICS");
  
  // Calculate summary stats
  const totalShows = artistShows.length;
  
  // Get first and last show dates
  const sortedShows = [...artistShows].sort((a, b) => {
    const dateA = a[dateCol] instanceof Date ? a[dateCol] : new Date(a[dateCol]);
    const dateB = b[dateCol] instanceof Date ? b[dateCol] : new Date(b[dateCol]);
    return dateA - dateB;
  });
  
  const firstShow = sortedShows[0][dateCol];
  const lastShow = sortedShows[sortedShows.length - 1][dateCol];
  
  // Format dates
  let firstShowStr = firstShow instanceof Date ? 
    Utilities.formatDate(firstShow, Session.getScriptTimeZone(), "MMM d, yyyy") : firstShow;
  
  let lastShowStr = lastShow instanceof Date ? 
    Utilities.formatDate(lastShow, Session.getScriptTimeZone(), "MMM d, yyyy") : lastShow;
  
  // Count unique venues and cities
  const venues = [...new Set(artistShows.map(show => show[venueCol]))];
  const cities = [...new Set(artistShows.map(show => show[cityCol]))];
  
  // Calculate average rating
  const ratings = artistShows.map(show => show[ratingCol])
    .filter(rating => rating !== "" && rating !== null && !isNaN(parseFloat(rating)));
  
  let avgRating = "N/A";
  if (ratings.length > 0) {
    const sum = ratings.reduce((acc, val) => acc + parseFloat(val), 0);
    avgRating = (sum / ratings.length).toFixed(1);
  }
  
  // Calculate time span
  let timeSpan = "N/A";
  try {
    if (firstShow instanceof Date && lastShow instanceof Date) {
      const diffTime = Math.abs(lastShow - firstShow);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      const diffYears = (diffDays / 365).toFixed(1);
      
      if (diffDays === 0) {
        timeSpan = "Same day";
      } else if (diffDays < 365) {
        timeSpan = `${diffDays} day${diffDays === 1 ? "" : "s"}`;
      } else {
        timeSpan = `${diffYears} year${diffYears === "1.0" ? "" : "s"}`;
      }
    }
  } catch (e) {
    timeSpan = "Error calculating";
  }
  
  // Create a visually appealing summary card
  const summaryValues = [
    ["Times Seen:", totalShows.toString()],
    ["First Saw:", firstShowStr],
    ["Last Saw:", lastShowStr],
    ["Time Span:", timeSpan],
    ["Unique Venues:", venues.length.toString()],
    ["Unique Cities:", cities.length.toString()],
    ["Average Rating:", avgRating]
  ];
  
  // Create summary table with two columns
  const leftCol = summaryValues.slice(0, 4);
  const rightCol = summaryValues.slice(4);
  
  // Set up two-column layout
  const maxRows = Math.max(leftCol.length, rightCol.length);
  
  // Set values for each column
  for (let i = 0; i < maxRows; i++) {
    // Left column
    if (i < leftCol.length) {
      artistSheet.getRange(row + i, 1).setValue(leftCol[i][0])
        .setFontWeight("bold")
        .setHorizontalAlignment("right")
        .setBackground(COLORS.PRIMARY_LIGHT);
      
      artistSheet.getRange(row + i, 2).setValue(leftCol[i][1])
        .setBackground(COLORS.SECONDARY);
    }
    
    // Right column
    if (i < rightCol.length) {
      artistSheet.getRange(row + i, 3).setValue(rightCol[i][0])
        .setFontWeight("bold")
        .setHorizontalAlignment("right")
        .setBackground(COLORS.PRIMARY_LIGHT);
      
      const valueCell = artistSheet.getRange(row + i, 4);
      valueCell.setValue(rightCol[i][1])
        .setBackground(COLORS.SECONDARY);
      
      // Special formatting for average rating - applying cell background color
      if (rightCol[i][0] === "Average Rating:" && rightCol[i][1] !== "N/A") {
        const ratingColor = getRatingColor(parseFloat(rightCol[i][1]));
        if (ratingColor) {
          valueCell.setBackground(ratingColor).setFontColor("white").setFontWeight("bold");
        }
      }
    }
  }
  
  row += maxRows + 2; // Add space after summary
  
  // SHOW LISTING - Now with Event Name and Headliner details
  row = createSectionHeader(artistSheet, row, "SHOW HISTORY");
  
  // Create header row - adding Event Name and adjusted columns
  let listingHeaders = ["Date", "Venue", "City", "Event Name", "Role/Headliner", "Rating", "Companions"];
  
  row = createTableHeader(artistSheet, row, listingHeaders);
  
  // Prepare for zebra striping
  let isAlternate = false;
  
  // Find artists who were headliners to identify them for support act rows
  // Create a map of events to their headliners
  const eventHeadliners = {};
  
  // First pass to identify headliners for each event
  data.slice(1).forEach(show => {
    // Skip non-attended shows
    if (show[statusCol] !== "Attended") return;
    
    const showArtist = show[artistCol];
    const eventName = show[eventNameCol] || "";
    const role = show[roleCol] || "";
    const showDate = show[dateCol];
    const venue = show[venueCol] || "";
    
    // Create a unique event key using date and venue
    let eventKey = "";
    if (showDate instanceof Date) {
      eventKey = `${Utilities.formatDate(showDate, Session.getScriptTimeZone(), "yyyy-MM-dd")}_${venue}`;
    } else if (typeof showDate === 'string') {
      eventKey = `${showDate}_${venue}`;
    }
    
    // Only create entries for headliners
    if (role.toLowerCase().includes("headline") || role.toLowerCase().includes("headliner")) {
      if (!eventHeadliners[eventKey]) {
        eventHeadliners[eventKey] = [];
      }
      // Add this artist as a headliner
      eventHeadliners[eventKey].push(showArtist);
    }
  });
  
  // Add each show with enhanced details
  sortedShows.forEach(show => {
    const showDate = show[dateCol] instanceof Date ? 
      Utilities.formatDate(show[dateCol], Session.getScriptTimeZone(), "MMM d, yyyy") : 
      show[dateCol];
    
    const rating = show[ratingCol] || "";
    const eventName = show[eventNameCol] || "";
    const role = show[roleCol] || "";
    
    // Create an event key to lookup headliners
    let eventKey = "";
    if (show[dateCol] instanceof Date) {
      eventKey = `${Utilities.formatDate(show[dateCol], Session.getScriptTimeZone(), "yyyy-MM-dd")}_${show[venueCol] || ""}`;
    } else if (typeof show[dateCol] === 'string') {
      eventKey = `${show[dateCol]}_${show[venueCol] || ""}`;
    }
    
    // Prepare the role/headliner display text
    let roleHeadlinerText = role;
    
    // Add headliner info for support acts
    if (role.toLowerCase().includes("support")) {
      const headliners = eventHeadliners[eventKey] || [];
      // Filter out the current artist (in case the role is incorrect)
      const otherHeadliners = headliners.filter(artist => artist !== artistName);
      
      if (otherHeadliners.length > 0) {
        roleHeadlinerText = `${role} for ${otherHeadliners.join(" & ")}`;
      }
    }
    
    // Prepare the row data with new columns
    let rowValues = [
      showDate,
      show[venueCol] || "",
      show[cityCol] || "",
      eventName,
      roleHeadlinerText,
      rating,
      show[withCol] || ""  // Add companions back to the main listing
    ];
    
    const rowRange = artistSheet.getRange(row, 1, 1, listingHeaders.length);
    rowRange.setValues([rowValues]);
    
    // Apply zebra striping
    rowRange.setBackground(isAlternate ? COLORS.SECONDARY : "white");
    isAlternate = !isAlternate;
    
    // Apply rating color to cell background if available
    if (rating && !isNaN(parseFloat(rating))) {
      const ratingColor = getRatingColor(parseFloat(rating));
      if (ratingColor) {
        artistSheet.getRange(row, 6).setBackground(ratingColor).setFontColor("white").setFontWeight("bold");
      }
    }
    
    row++;
  });
  
  row += 1; // Space after listing
  
  // Adjust column widths to accommodate the new data
  artistSheet.setColumnWidth(4, 180); // Event Name column wider
  artistSheet.setColumnWidth(5, 180); // Role/Headliner column wider
  artistSheet.setColumnWidth(7, 150); // Companions column wider
  
  // VENUE BREAKDOWN
  row = createSectionHeader(artistSheet, row, "VENUES");
  
  // Count occurrences of each venue
  const venueCounts = {};
  artistShows.forEach(show => {
    const venue = show[venueCol] || "Unknown";
    venueCounts[venue] = (venueCounts[venue] || 0) + 1;
  });
  
  // Sort venues by frequency
  const sortedVenues = Object.entries(venueCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([venue, count]) => [venue, count]);
  
  // Add venue headers
  row = createTableHeader(artistSheet, row, ["Venue", "Times"]);
  
  // Add venue data with zebra striping
  if (sortedVenues.length > 0) {
    const venueRange = artistSheet.getRange(row, 1, sortedVenues.length, 2);
    venueRange.setValues(sortedVenues);
    venueRange.setBackgrounds(sortedVenues.map((_, i) => [
      i % 2 === 1 ? COLORS.SECONDARY : "white",
      i % 2 === 1 ? COLORS.SECONDARY : "white"
    ]));
    artistSheet.getRange(row, 2, sortedVenues.length, 1)
      .setFontWeights(sortedVenues.map(() => ["bold"]));
    row += sortedVenues.length + 1;
  } else {
    artistSheet.getRange(row, 1, 1, 2).merge()
      .setValue("No venue data available")
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
    row += 2;
  }
  
  // CITY BREAKDOWN
  row = createSectionHeader(artistSheet, row, "CITIES");
  
  // Count occurrences of each city
  const cityCounts = {};
  artistShows.forEach(show => {
    const city = show[cityCol] || "Unknown";
    cityCounts[city] = (cityCounts[city] || 0) + 1;
  });
  
  // Sort cities by frequency
  const sortedCities = Object.entries(cityCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([city, count]) => [city, count]);
  
  // Add city headers
  row = createTableHeader(artistSheet, row, ["City", "Times"]);
  
  // Add city data with zebra striping
  if (sortedCities.length > 0) {
    const cityRange = artistSheet.getRange(row, 1, sortedCities.length, 2);
    cityRange.setValues(sortedCities);
    cityRange.setBackgrounds(sortedCities.map((_, i) => [
      i % 2 === 1 ? COLORS.SECONDARY : "white",
      i % 2 === 1 ? COLORS.SECONDARY : "white"
    ]));
    artistSheet.getRange(row, 2, sortedCities.length, 1)
      .setFontWeights(sortedCities.map(() => ["bold"]));
    row += sortedCities.length + 1;
  } else {
    artistSheet.getRange(row, 1, 1, 2).merge()
      .setValue("No city data available")
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
    row += 2;
  }
  
  // COMPANION BREAKDOWN
  row = createSectionHeader(artistSheet, row, "CONCERT COMPANIONS");
  
  // Count companions
  const companionCounts = {};
  artistShows.forEach(show => {
    if (show[withCol]) {
      const companions = show[withCol].split(',').map(c => c.trim());
      companions.forEach(companion => {
        if (companion) {
          companionCounts[companion] = (companionCounts[companion] || 0) + 1;
        }
      });
    }
  });
  
  // Sort companions by frequency
  const sortedCompanions = Object.entries(companionCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([companion, count]) => [companion, count]);
  
  // Add companion headers
  row = createTableHeader(artistSheet, row, ["Companion", "Times"]);
  
  // Add companion data with zebra striping
  if (sortedCompanions.length > 0) {
    const companionRange = artistSheet.getRange(row, 1, sortedCompanions.length, 2);
    companionRange.setValues(sortedCompanions);
    companionRange.setBackgrounds(sortedCompanions.map((_, i) => [
      i % 2 === 1 ? COLORS.SECONDARY : "white",
      i % 2 === 1 ? COLORS.SECONDARY : "white"
    ]));
    artistSheet.getRange(row, 2, sortedCompanions.length, 1)
      .setFontWeights(sortedCompanions.map(() => ["bold"]));
    row += sortedCompanions.length + 1;
  } else {
    artistSheet.getRange(row, 1, 1, 2).merge()
      .setValue("No companion data available")
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
    row += 2;
  }
  
  // SETLIST ANALYSIS (if setlist URLs are available)
  if (setlistCol !== -1) {
    // Check if any shows have setlist URLs
    const hasSetlists = artistShows.some(show => show[setlistCol] && show[setlistCol] !== "");
    
    // Debug logs to help diagnose setlist issues
    Logger.log(`Checking for setlists: setlistCol=${setlistCol}, hasSetlists=${hasSetlists}`);
    if (hasSetlists) {
      Logger.log("Found shows with setlist URLs:");
      artistShows.filter(show => show[setlistCol] && show[setlistCol] !== "")
                .forEach((show, i) => Logger.log(`  [${i+1}] ${show[setlistCol]}`));
    }
    
    // Create setlist analysis section
    row = createSectionHeader(artistSheet, row, "SETLIST ANALYSIS");
    
    // Try to integrate setlist data if function exists
    try {
      if (typeof integrateSetlistData === 'function') {
        row = integrateSetlistData(artistSheet, artistShows, row, headers);
      } else {
        // Since the integration function doesn't exist, let's implement a simple version
        // that shows the top 20 songs from setlists
        
        // Create a section for top songs
        row = createSectionHeader(artistSheet, row, "TOP 20 SONGS");
        
        // Check if we have a Songs column (try to find it)
        const songsCol = headers.indexOf("Songs");
        
        if (songsCol !== -1) {
          // Count song occurrences across all shows
          const songCounts = {};
          artistShows.forEach(show => {
            if (show[songsCol]) {
              const songs = show[songsCol].split(',').map(s => s.trim());
              songs.forEach(song => {
                if (song) {
                  songCounts[song] = (songCounts[song] || 0) + 1;
                }
              });
            }
          });
          
          // Sort songs by frequency and take top 20
          const sortedSongs = Object.entries(songCounts)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 20)
            .map(([song, count]) => [song, count]);
          
          // Add column headers for songs
          row = createTableHeader(artistSheet, row, ["Song", "Times Played"]);
          
          // Add song data with zebra striping
          let isAlternate = false;
          if (sortedSongs.length > 0) {
            sortedSongs.forEach(song => {
              const songRow = artistSheet.getRange(row, 1, 1, 2);
              songRow.setValues([song])
                .setBackground(isAlternate ? COLORS.SECONDARY : "white");
              
              // Apply bold to frequency number
              artistSheet.getRange(row, 2).setFontWeight("bold");
              
              isAlternate = !isAlternate;
              row++;
            });
            row += 1;
          } else {
            artistSheet.getRange(row, 1, 1, 2).merge()
              .setValue("No song data available")
              .setFontStyle("italic")
              .setHorizontalAlignment("center");
            row += 2;
          }
        } else {
          // No Songs column found, show placeholder message
          const noSongsRange = artistSheet.getRange(row, 1, 1, 6);
          noSongsRange.merge()
            .setValue("No song data available. Add a 'Songs' column to your tracker to enable this feature.")
            .setFontStyle("italic")
            .setBackground(COLORS.SECONDARY)
            .setHorizontalAlignment("center");
          row += 2;
          
          // Add a suggestion for implementation
          const suggestRange = artistSheet.getRange(row, 1, 1, 6);
          suggestRange.merge()
            .setValue("Tip: For each concert, add comma-separated song names in the 'Songs' column")
            .setFontStyle("italic")
            .setFontColor(COLORS.TEXT_SECONDARY)
            .setHorizontalAlignment("center");
          row += 3;
        }
        
        // Add placeholder text with styled formatting for full setlist integration
        const noSetlistRange = artistSheet.getRange(row, 1, 1, 6);
        noSetlistRange.merge()
          .setValue("For additional setlist analysis, install the setlist-integration.js module.")
          .setFontStyle("italic")
          .setBackground(COLORS.SECONDARY)
          .setHorizontalAlignment("center");
        row += 2;
      }
    } catch (error) {
      // Log any errors that occur during setlist integration
      Logger.log(`Error in setlist integration: ${error.toString()}`);
      
      // Add error message with styled formatting
      const errorRange = artistSheet.getRange(row, 1, 1, 6);
      errorRange.merge()
        .setValue(`Error processing setlist data: ${error.toString()}`)
        .setFontColor("#D32F2F") // Error red
        .setBackground("#FFEBEE") // Light red background
        .setHorizontalAlignment("center");
      row += 3;
    }
  }
  
  // Add timestamp and footer - more compact
  const now = new Date();
  const footerRange = artistSheet.getRange(row, 1, 1, 7);
  footerRange.merge()
    .setValue(`Generated: ${Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM d, yyyy h:mm a")}`)
    .setFontStyle("italic")
    .setFontColor(COLORS.TEXT_SECONDARY)
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_LIGHT);
  
  // Auto-resize columns for best display
  artistSheet.autoResizeColumns(1, 6);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`Analysis for ${artistName} complete!`, "Success");
}

/**
 * Handle the button click by checking if the clicked cell is the analyse button
 */
function handleArtistBreakdownClick(e) {
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  if (sheet.getName() !== ARTIST_SHEET_NAME) return;
  
  // Check if the clicked cell is C5 (the analyse button)
  if (e.range.getA1Notation() === "C5") {
    analyzeSelectedArtist();
  }
}

/**
 * Add an event trigger to handle clicks on the sheet
 */
function setUpArtistBreakdownTrigger() {
  // Clean up existing triggers
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'handleArtistBreakdownClick') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
  
  // Create a new trigger
  ScriptApp.newTrigger('handleArtistBreakdownClick')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Artist Breakdown button trigger created", "Setup Complete");
}

/**
 * Update menu to include Artist Breakdown
 */
function updateConcertTrackerMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Concert Tracker')
    .addItem('Generate Calendar Views', 'generateCalendarViews')
    .addItem('Update Analysis Dashboard', 'updateAnalysis')
    .addSeparator()
    .addItem('Artist Analysis Dashboard', 'createArtistBreakdown')
    .addToUi();
}

/**
 * Run this to manually set up the Artist Breakdown dashboard
 */
function initializeArtistBreakdown() {
  createArtistBreakdown();
  setUpArtistBreakdownTrigger();
  SpreadsheetApp.getActiveSpreadsheet().toast("Artist Breakdown dashboard initialised", "Setup Complete");
}