/**
 * Create a chart for year-based data with blue-based color scheme
 */
function createYearChart(sheet, startRow, dataRows, title) {
  // Check if dataRows is defined and has the expected structure
  if (!dataRows || !Array.isArray(dataRows) || dataRows.length <= 2) {
    console.log("Not enough data for chart or invalid data format");
    return; // Exit without creating chart
  }
  
  try {
    // Extract data for chart (skip header row)
    const years = [];
    const counts = [];
    
    for (let i = 1; i < dataRows.length; i++) {
      if (dataRows[i] && dataRows[i][0] !== undefined) {
        years.push(dataRows[i][0].toString());
        counts.push(dataRows[i][1]);
      }
    }
    
    // Verify we have data to chart
    if (years.length === 0 || counts.length === 0) {
      console.log("No valid data points found for chart");
      return;
    }
    
    // Create chart data range
    const chartDataRange = sheet.getRange(startRow + 1, 4, years.length + 1, 2);
    const chartData = [
      ["Year", "Count"]
    ];
    
    // Add data rows
    for (let i = 0; i < years.length; i++) {
      chartData.push([years[i], counts[i]]);
    }
    
    // Set the values in the sheet
    chartDataRange.setValues(chartData);
    
    // Create the chart with blue-based styling
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(chartDataRange)
      .setPosition(startRow + 2, 4, 0, 0)
      .setOption('title', title)
      .setOption('titleTextStyle', {fontSize: 14, bold: true})
      .setOption('hAxis.title', 'Year')
      .setOption('vAxis.title', title.includes('Artist') ? 'Unique Artists' : 'Events')
      .setOption('legend', {position: 'none'})
      .setOption('colors', [ANALYSIS_COLORS.CHART_COLOR]) // Steel Blue from constants
      .setOption('width', 500)
      .setOption('height', 300)
      .build();
    
    sheet.insertChart(chart);
  } catch (e) {
    console.error("Error creating chart:", e);
    // Continue execution without failing the entire function
  }
}

/**
 * Convert an object of counts to sorted rows for display
 * Modified to ensure chronological display for years without showing zero counts
 */
function objectToSortedRows(countObject, labelName, countName, chronological = false) {
  // Safety check
  if (!countObject || Object.keys(countObject).length === 0) {
    return [[labelName, countName]]; // Return just the header row
  }
  
  try {
    // Convert object to array of [key, value] pairs
    let pairs = Object.entries(countObject);
    
    // Sort by count (descending) or chronologically if specified
    if (chronological) {
      if (labelName === "Year") {
        // For year data, ensure proper chronological sorting
        // DON'T include zero count years, but ensure correct chronological order
        pairs.sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
      } else {
        // For non-year data, just sort chronologically
        pairs.sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
      }
    } else {
      // First sort alphabetically by key (secondary sort)
      pairs.sort((a, b) => a[0].toString().localeCompare(b[0].toString()));
      // Then sort by count (primary sort, descending)
      pairs.sort((a, b) => b[1] - a[1]);
    }
    
    // Format as rows with headers
    const rows = [
      [labelName, countName],
      ...pairs.map(pair => [pair[0], pair[1]])
    ];
    
    return rows;
  } catch (e) {
    console.error("Error in objectToSortedRows:", e);
    return [[labelName, countName]]; // Return just the header row in case of error
  }
}

/**
 * Concert Tracker Google Sheets System
 * 
 * This script creates a comprehensive concert tracking system with:
 * - Real calendar view of upcoming/announced shows
 * - Mobile-friendly list with formatted dates (Wed 13 Aug)
 * - Analysis dashboard of attended concerts
 */

// Global variables
const SHEET_NAMES = {
  MAIN: "Tracker",
  CALENDAR: "Calendar View",
  MOBILE_CALENDAR: "Mobile Calendar",
  ANALYSIS: "Concert Analysis",
};

// Define color schemes
const STATUS_COLORS = {
  UPCOMING: "#20B2AA", // Teal
  ANNOUNCED: "#FFA500" // Orange
};

// Define blue color scheme constants with a different name to avoid conflicts
const ANALYSIS_COLORS = {
  TITLE_BG: "#1e88e5",        // Main title background (blue)
  TITLE_TEXT: "#FFFFFF",      // Main title text (white)
  HEADER_BG: "#0d47a1",       // Section headers background (darker blue)
  HEADER_TEXT: "#FFFFFF",     // Section headers text (white)
  CHART_COLOR: "#4682B4",     // SteelBlue for charts
  ALT_ROW_COLOR: "#F0F8FF",   // AliceBlue for alternating rows (optional)
  HIGHLIGHT: "#1E90FF"        // DodgerBlue for highlights (optional)
};

const DAYS_OF_WEEK = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/**
 * Update the analysis dashboard with statistics from attended concerts
 * Improved to show years with proper spacing and filtering concert companions
 */
function updateAnalysis() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create or get sheets
    let mainSheet = ss.getSheetByName(SHEET_NAMES.MAIN);
    if (!mainSheet) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Error: Main tracker sheet not found", "Error");
      return;
    }
    
    let analysisSheet = getOrCreateSheet(SHEET_NAMES.ANALYSIS);
    
    // Clear existing analysis data
    analysisSheet.clear();
    
    // Get all concert data
    const data = mainSheet.getDataRange().getValues();
    if (!data || data.length <= 1) {
      // Only header row or no data
      analysisSheet.getRange("A1").setValue("Concert Analysis Dashboard").setFontWeight("bold").setFontSize(16)
      .setFontColor(ANALYSIS_COLORS.CHART_COLOR); // Added blue color to title
      analysisSheet.getRange("A3").setValue("No concert data found to analyze.").setFontStyle("italic");
      return;
    }
    
    const headers = data[0];
    
    // Find column indexes (with safety checks)
    const artistCol = headers.indexOf("Artist");
    const dateCol = headers.indexOf("Date");
    const venueCol = headers.indexOf("Venue");
    const cityCol = headers.indexOf("City");
    const withCol = headers.indexOf("Who I Went With");
    const statusCol = headers.indexOf("Status");
    const ratingCol = headers.indexOf("Rating");
    
    // Check that required columns exist
    if (artistCol === -1 || dateCol === -1 || statusCol === -1) {
      analysisSheet.getRange("A1").setValue("Concert Analysis Dashboard").setFontWeight("bold").setFontSize(16);
      analysisSheet.getRange("A3").setValue("Required columns missing (Artist, Date, or Status)").setFontStyle("italic");
      return;
    }
    
    // Filter for attended shows
    const attendedShows = data.slice(1).filter(row => row[statusCol] === "Attended");
    
    if (attendedShows.length === 0) {
      // No attended shows found
      analysisSheet.getRange("A1").setValue("Concert Analysis Dashboard").setFontWeight("bold").setFontSize(16);
      analysisSheet.getRange("A3").setValue("No attended concerts found to analyze.").setFontStyle("italic");
      return;
    }
    
    // Group by events (same date + venue = same event)
    const events = {};
    const eventsByVenue = {};
    const eventsByCity = {};
    const eventsByYear = {};
    const eventsByCompanion = {};
    
    // Artist counter (independent of events)
    const artistCounts = {};
    const artistsByYear = {};
    
    // Process all attended shows
    attendedShows.forEach(show => {
      // Artist count (always individual)
      const artist = show[artistCol];
      if (artist) {
        artistCounts[artist] = (artistCounts[artist] || 0) + 1;
      }
      
      // Format date for event key
      let dateStr = "";
      let yearVal = "";
      
      if (show[dateCol] instanceof Date) {
        dateStr = Utilities.formatDate(show[dateCol], Session.getScriptTimeZone(), "yyyy-MM-dd");
        yearVal = show[dateCol].getFullYear();
      } else if (typeof show[dateCol] === "string") {
        // Try to parse the date string
        let dateObj;
        if (show[dateCol].includes('/')) {
          const dateParts = show[dateCol].split('/');
          if (dateParts.length === 3) {
            // Try to handle dd/mm/yyyy format
            const day = parseInt(dateParts[0], 10);
            const month = parseInt(dateParts[1], 10) - 1;
            const year = parseInt(dateParts[2], 10);
            dateObj = new Date(year, month, day);
          }
        } else {
          dateObj = new Date(show[dateCol]);
        }
        
        if (!isNaN(dateObj.getTime())) {
          dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
          yearVal = dateObj.getFullYear();
        } else if (show[dateCol].trim().match(/^\d{4}$/)) {
          // Year only format
          dateStr = show[dateCol].trim() + "-01-01"; // Use Jan 1 as placeholder
          yearVal = parseInt(show[dateCol].trim());
        }
      }
      
      // Track artists by year
      if (yearVal && artist) {
        if (!artistsByYear[yearVal]) {
          artistsByYear[yearVal] = new Set();
        }
        artistsByYear[yearVal].add(artist);
      }
      
      // Create an event key (date + venue)
      const venue = show[venueCol] || "";
      const eventKey = `${dateStr}|${venue}`;
      
      // Count event only once
      if (!events[eventKey]) {
        events[eventKey] = {
          date: dateStr,
          venue: venue,
          city: show[cityCol] || "",
          year: yearVal,
          artists: [],
          companions: []
        };
      }
      
      // Add this artist to the event
      if (artist) {
        events[eventKey].artists.push(artist);
      }
      
      // Add companions to the event (avoiding duplicates)
      if (withCol >= 0 && show[withCol]) {
        const companions = show[withCol].split(",").map(c => c.trim());
        companions.forEach(companion => {
          if (companion && !events[eventKey].companions.includes(companion)) {
            events[eventKey].companions.push(companion);
          }
        });
      }
    });
    
    // Count events by various dimensions
    Object.values(events).forEach(event => {
      // Venue count
      if (event.venue) {
        eventsByVenue[event.venue] = (eventsByVenue[event.venue] || 0) + 1;
      }
      
      // City count
      if (event.city) {
        eventsByCity[event.city] = (eventsByCity[event.city] || 0) + 1;
      }
      
      // Year count
      if (event.year) {
        eventsByYear[event.year] = (eventsByYear[event.year] || 0) + 1;
      }
      
      // Count companions by event
      event.companions.forEach(companion => {
        eventsByCompanion[companion] = (eventsByCompanion[companion] || 0) + 1;
      });
    });
    
    // Calculate average rating
    let avgRating = "N/A";
    if (ratingCol >= 0) {
      const ratings = attendedShows.map(show => show[ratingCol])
        .filter(rating => rating !== "" && rating !== null && !isNaN(parseFloat(rating)));
      
      if (ratings.length > 0) {
        const sum = ratings.reduce((acc, val) => acc + parseFloat(val), 0);
        avgRating = (sum / ratings.length).toFixed(1);
      }
    }
    
    // Format the analysis sheet - matching Artist Breakdown styling
    analysisSheet.getRange("A1").setValue("Concert Analysis Dashboard").setFontWeight("bold").setFontSize(16);
    
    // SUMMARY STATISTICS - styled like Artist Breakdown
    let currentRow = 3;
    analysisSheet.getRange(currentRow, 1).setValue("Summary Statistics").setFontWeight("bold").setFontSize(14)
      .setFontColor(ANALYSIS_COLORS.CHART_COLOR); // Added blue color to section header
    currentRow += 1;
    
    const summaryStats = [
      ["Total Events Attended:", Object.keys(events).length],
      ["Unique Artists Seen:", Object.keys(artistCounts).length],
      ["Unique Venues Visited:", Object.keys(eventsByVenue).length],
      ["Unique Cities Visited:", Object.keys(eventsByCity).length],
      ["Average Rating:", avgRating]
    ];
    
    analysisSheet.getRange(currentRow, 1, summaryStats.length, 2).setValues(summaryStats);
    analysisSheet.getRange(currentRow, 1, summaryStats.length, 1).setFontWeight("bold");
    currentRow += summaryStats.length + 2; // Add some spacing
    
    // MOST FREQUENT ARTISTS
    const artistHeaderRange = analysisSheet.getRange(currentRow, 1, 1, 3);
    artistHeaderRange.merge()
      .setValue("Most Seen Artists")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground(ANALYSIS_COLORS.HEADER_BG)
      .setFontColor(ANALYSIS_COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");
    currentRow += 1;
    
    // Format table headers with blue background and white text
    const artistRows = objectToSortedRows(artistCounts, "Artist", "Times Seen");
    analysisSheet.getRange(currentRow, 1, 1, 2).setValues([artistRows[0]])
      .setFontWeight("bold")
      .setBackground(ANALYSIS_COLORS.HEADER_BG)
      .setFontColor(ANALYSIS_COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");
    currentRow += 1;
    
    // Add artist data (only include artists with count >= 2)
    const filteredArtistData = artistRows.slice(1).filter(row => parseInt(row[1]) >= 2);
    if (filteredArtistData.length > 0) {
      analysisSheet.getRange(currentRow, 1, filteredArtistData.length, 2).setValues(filteredArtistData);
      currentRow += filteredArtistData.length + 2;
    } else {
      analysisSheet.getRange(currentRow, 1).setValue("No artists seen multiple times");
      currentRow += 2;
    }
    
    // TOP VENUES
    const venueHeaderRange = analysisSheet.getRange(currentRow, 1, 1, 3);
    venueHeaderRange.merge()
      .setValue("Most Visited Venues")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground(ANALYSIS_COLORS.HEADER_BG)
      .setFontColor(ANALYSIS_COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");
    currentRow += 1;
    
    const venueRows = objectToSortedRows(eventsByVenue, "Venue", "Events");
    analysisSheet.getRange(currentRow, 1, 1, 2).setValues([venueRows[0]])
      .setFontWeight("bold").setBackground("#B0C4DE").setHorizontalAlignment("center");
    currentRow += 1;
    
    // Add venue data (only include venues with count >= 2)
    const filteredVenueData = venueRows.slice(1).filter(row => parseInt(row[1]) >= 2);
    if (filteredVenueData.length > 0) {
      analysisSheet.getRange(currentRow, 1, filteredVenueData.length, 2).setValues(filteredVenueData);
      currentRow += filteredVenueData.length + 2;
    } else {
      analysisSheet.getRange(currentRow, 1).setValue("No venues visited multiple times");
      currentRow += 2;
    }
    
    // TOP CITIES
    const cityHeaderRange = analysisSheet.getRange(currentRow, 1, 1, 3);
    cityHeaderRange.merge()
      .setValue("Most Visited Cities")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground(ANALYSIS_COLORS.HEADER_BG)
      .setFontColor(ANALYSIS_COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");
    currentRow += 1;
    
    const cityRows = objectToSortedRows(eventsByCity, "City", "Events");
    analysisSheet.getRange(currentRow, 1, 1, 2).setValues([cityRows[0]])
      .setFontWeight("bold").setBackground("#B0C4DE").setHorizontalAlignment("center");
    currentRow += 1;
    
    // Add city data
    if (cityRows.length > 1) {
      analysisSheet.getRange(currentRow, 1, cityRows.length - 1, 2).setValues(cityRows.slice(1));
      currentRow += cityRows.length + 2;
    } else {
      analysisSheet.getRange(currentRow, 1).setValue("No city data available");
      currentRow += 2;
    }
    
    // ===== EVENTS BY YEAR - with chart =====
    const eventsYearStartRow = currentRow;
    const eventsYearHeaderRange = analysisSheet.getRange(currentRow, 1, 1, 3);
    eventsYearHeaderRange.merge()
      .setValue("Events by Year")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground(ANALYSIS_COLORS.HEADER_BG)
      .setFontColor(ANALYSIS_COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");
    currentRow += 1;
    
    // Process events by year data - only show years with data, no zero counts
    processYearData(analysisSheet, currentRow, eventsByYear, "Events", eventsYearStartRow);
    
    // Make sure we have enough room for the chart, calculate the chart size
    const minRows = 12; // Minimum rows to reserve for the chart
    const yearCount = Object.keys(eventsByYear).length;
    const rowsNeeded = Math.max(minRows, yearCount + 5); // Add 5 for headers and spacing
    
    // Update the currentRow after processing year data
    currentRow = eventsYearStartRow + rowsNeeded;
    
    // ===== ARTISTS BY YEAR - with chart =====
    const artistsYearStartRow = currentRow;
    const artistsYearHeaderRange = analysisSheet.getRange(currentRow, 1, 1, 3);
    artistsYearHeaderRange.merge()
      .setValue("Unique Artists by Year")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground(ANALYSIS_COLORS.HEADER_BG)
      .setFontColor(ANALYSIS_COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");
    currentRow += 1;
    
    // Process artists by year data - only show years with data, no zero counts
    const artistYearData = {};
    try {
      Object.keys(artistsByYear).sort().forEach(year => {
        if (artistsByYear[year] instanceof Set) {
          artistYearData[year] = artistsByYear[year].size;
        } else if (Array.isArray(artistsByYear[year])) {
          artistYearData[year] = new Set(artistsByYear[year]).size;
        }
      });
    } catch (e) {
      console.error("Error processing artist year data:", e);
    }
    
    processYearData(analysisSheet, currentRow, artistYearData, "Unique Artists", artistsYearStartRow);
    
    // Make sure we have enough room for the chart, calculate the chart size
    const minArtistRows = 12; // Minimum rows to reserve for the chart
    const artistYearCount = Object.keys(artistYearData).length;
    const artistRowsNeeded = Math.max(minArtistRows, artistYearCount + 5); // Add 5 for headers and spacing
    
    // Update the currentRow after processing artist year data
    currentRow = artistsYearStartRow + artistRowsNeeded;
    
    // ===== CONCERT COMPANIONS - Filter to only show companions with count >= 2 =====
    const companionsHeaderRange = analysisSheet.getRange(currentRow, 1, 1, 3);
    companionsHeaderRange.merge()
      .setValue("Concert Companions")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground(ANALYSIS_COLORS.HEADER_BG)
      .setFontColor(ANALYSIS_COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");
    currentRow += 1;
    
    // Process companion data
    const companionRows = objectToSortedRows(eventsByCompanion, "Companion", "Times");
    
    // Create header with styling
    analysisSheet.getRange(currentRow, 1, 1, 2).setValues([companionRows[0]])
      .setFontWeight("bold").setBackground("#E0E0E0").setHorizontalAlignment("center");
    currentRow += 1;
    
    // Add companion data rows - ONLY THOSE WITH COUNT >= 2
    const filteredCompanions = companionRows.slice(1).filter(row => parseInt(row[1]) >= 2);
    if (filteredCompanions.length > 0) {
      analysisSheet.getRange(currentRow, 1, filteredCompanions.length, 2).setValues(filteredCompanions);
      currentRow += filteredCompanions.length;
    } else {
      analysisSheet.getRange(currentRow, 1).setValue("No companions with 2 or more events together");
      currentRow += 2;
    }
    
    // Add timestamp
    currentRow += 2;
    const now = new Date();
    // Match timestamp format from Artist Breakdown
    analysisSheet.getRange(currentRow, 1).setValue(`Analysis generated: ${Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM d, yyyy h:mm a")}`)
      .setFontStyle("italic");
    
    // Format and finalize
    analysisSheet.setColumnWidth(1, 200);
    analysisSheet.setColumnWidth(2, 200);
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Analysis dashboard updated successfully!", "Success");
  } catch (error) {
    console.error("Error updating analysis dashboard:", error);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error updating analysis. See logs for details.", "Error");
  }
}

/**
 * Helper function to process and display year-based data
 * This function handles both Events by Year and Artists by Year data
 * Modified to ensure proper spacing even if some years have no data
 */
function processYearData(sheet, startRow, dataObject, label, chartStartRow) {
  try {
    // Create sorted rows for year data (chronological order)
    // Only includes years with actual data (no zero counts)
    const yearRows = objectToSortedRows(dataObject, "Year", label, true);
    
    // Add the header row with styling using blue color scheme
    sheet.getRange(startRow, 1, 1, 2).setValues([yearRows[0]])
      .setFontWeight("bold")
      .setBackground(ANALYSIS_COLORS.HEADER_BG)
      .setFontColor(ANALYSIS_COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");
    
    // Add the data rows
    if (yearRows.length > 1) {
      sheet.getRange(startRow + 1, 1, yearRows.length - 1, 2).setValues(yearRows.slice(1));
      
      // Create chart with all years (for visual continuity)
      if (yearRows.length > 2) {
        // For chart display, figure out min/max years to ensure proper display
        const years = yearRows.slice(1).map(row => parseInt(row[0]));
        const minYear = Math.min(...years);
        const maxYear = Math.max(...years);
        
        createYearChart(sheet, chartStartRow, yearRows, `${label} by Year`);
      }
    } else {
      sheet.getRange(startRow + 1, 1).setValue(`No ${label.toLowerCase()} by year data available`);
    }
  } catch (e) {
    console.error(`Error processing ${label} by year:`, e);
    sheet.getRange(startRow + 1, 1).setValue(`Error processing ${label.toLowerCase()} by year data`);
  }
}

/**
 * Optimized function to generate all calendar views more efficiently
 * This replaces the existing generateCalendarViews function
 */
function generateCalendarViews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Start a timer to measure performance
  const startTime = new Date().getTime();
  
  // Create or get sheets - do this once at the beginning
  let mainSheet = ss.getSheetByName(SHEET_NAMES.MAIN);
  let calendarSheet = getOrCreateSheet(SHEET_NAMES.CALENDAR);
  let mobileCalendarSheet = getOrCreateSheet(SHEET_NAMES.MOBILE_CALENDAR);
  let upcomingMobileSheet = getOrCreateSheet("Upcoming Mobile Calendar");
  
  // Clear all sheets at once
  clearSheets([calendarSheet, mobileCalendarSheet, upcomingMobileSheet]);
  
  // Get all concert data at once
  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find all needed column indexes in one pass
  const columns = getColumnIndexes(headers, [
    "Artist", "Date", "Venue", "City", "Status", "Ticket Vendor"
  ]);
  
  // Pre-process all shows at once - parse dates and filter in one operation
  const { allUpcomingShows, onlyUpcomingShows } = preprocessShows(data, columns);
  
  // If no shows, display message and exit early
  if (allUpcomingShows.length === 0) {
    displayNoShowsMessage(calendarSheet, mobileCalendarSheet, upcomingMobileSheet);
    return;
  }
  
  // Calculate date ranges for all views at once
  const dateInfo = calculateDateRanges(allUpcomingShows);
  
  // Set up sheet headers in parallel
  setUpSheetHeaders(calendarSheet, mobileCalendarSheet, upcomingMobileSheet, allUpcomingShows.length, onlyUpcomingShows.length);
  
  // Process the desktop calendar view
  processDesktopCalendar(calendarSheet, allUpcomingShows, dateInfo, columns);
  
  // Process the mobile month-by-month view (all shows)
  processMobileCalendar(mobileCalendarSheet, allUpcomingShows, columns);
  
  // Process the upcoming-only mobile view
  processUpcomingMobileView(upcomingMobileSheet, onlyUpcomingShows, columns);
  
  // Format all sheets in parallel using batch operations
  formatSheets(calendarSheet, mobileCalendarSheet, upcomingMobileSheet);
  
  // Calculate and show execution time
  const endTime = new Date().getTime();
  const executionTime = (endTime - startTime) / 1000;
  SpreadsheetApp.getActiveSpreadsheet().toast(`All calendar views updated successfully in ${executionTime.toFixed(2)} seconds!`);
}

/**
 * Get all needed column indexes in one pass
 */
function getColumnIndexes(headers, columnNames) {
  const result = {};
  columnNames.forEach(name => {
    result[name.toLowerCase().replace(/\s+/g, '')] = headers.indexOf(name);
  });
  return result;
}

/**
 * Clear multiple sheets at once
 */
function clearSheets(sheets) {
  sheets.forEach(sheet => sheet.clear());
}

/**
 * Parse dates and filter shows in one pass
 */
function preprocessShows(data, columns) {
  const artistCol = columns.artist;
  const dateCol = columns.date;
  const venueCol = columns.venue;
  const statusCol = columns.status;
  
  // Process all shows at once
  const processedShows = [];
  
  // Start from row 1 (skip headers)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusCol];
    
    // Only process upcoming/announced shows
    if (status !== "Upcoming" && status !== "Announced") continue;
    
    // Parse date once
    let showDate = row[dateCol];
    if (!(showDate instanceof Date)) {
      if (typeof showDate === 'string' && showDate.includes('/')) {
        try {
          const dateParts = showDate.split('/');
          if (dateParts.length === 3) {
            const day = parseInt(dateParts[0], 10);
            const month = parseInt(dateParts[1], 10) - 1;
            const year = parseInt(dateParts[2], 10);
            showDate = new Date(year, month, day);
          }
        } catch (e) {
          // Invalid date, skip this show
          continue;
        }
      } else {
        // Not a valid date format, skip this show
        continue;
      }
    }
    
    // Skip invalid dates
    if (!(showDate instanceof Date) || isNaN(showDate.getTime())) continue;
    
    // Add processed show with pre-calculated values for efficiency
    processedShows.push({
      original: row,
      parsedDate: showDate,
      status: status,
      artist: row[artistCol],
      venue: row[venueCol],
      // Pre-compute date parts for calendar sorting
      year: showDate.getFullYear(),
      month: showDate.getMonth(),
      day: showDate.getDate(),
      dayOfWeek: showDate.getDay()
    });
  }
  
  // Sort all shows by date once
  processedShows.sort((a, b) => a.parsedDate - b.parsedDate);
  
  // Split into two arrays - one for all shows, one for only upcoming
  return {
    allUpcomingShows: processedShows,
    onlyUpcomingShows: processedShows.filter(show => show.status === "Upcoming")
  };
}

/**
 * Calculate all date ranges in one pass
 */
function calculateDateRanges(shows) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Find max date in one pass
  let maxDate = new Date(today);
  shows.forEach(show => {
    if (show.parsedDate > maxDate) {
      maxDate = new Date(show.parsedDate);
    }
  });
  
  const currentMonth = today.getMonth() + today.getFullYear() * 12;
  const maxMonth = maxDate.getMonth() + maxDate.getFullYear() * 12;
  const monthsToDisplay = Math.max(1, maxMonth - currentMonth + 1);
  
  // Generate month info for all needed months
  const monthsInfo = [];
  for (let i = 0; i < monthsToDisplay; i++) {
    const monthDate = new Date(today);
    monthDate.setMonth(today.getMonth() + i);
    
    const year = monthDate.getFullYear();
    const month = monthDate.getMonth();
    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    monthsInfo.push({
      date: monthDate,
      year: year,
      month: month,
      firstDay: firstDay,
      daysInMonth: daysInMonth,
      monthName: MONTHS[month],
      monthYear: `${MONTHS[month]} ${year}`
    });
  }
  
  return {
    today: today,
    monthsInfo: monthsInfo,
    monthsToDisplay: monthsToDisplay
  };
}

/**
 * Display message when no shows are found
 */
function displayNoShowsMessage(calendarSheet, mobileCalendarSheet, upcomingMobileSheet) {
  // Set up titles
  calendarSheet.getRange("A1").setValue("Concert Calendar").setFontWeight("bold").setFontSize(16);
  mobileCalendarSheet.getRange("A1").setValue("Mobile Concert Calendar").setFontWeight("bold").setFontSize(14);
  upcomingMobileSheet.getRange("A1").setValue("Upcoming Shows Only").setFontWeight("bold").setFontSize(16);
  
  // Set message
  calendarSheet.getRange("A3").setValue("No upcoming or announced shows found.").setFontStyle("italic");
  mobileCalendarSheet.getRange("A3").setValue("No upcoming or announced shows found.").setFontStyle("italic");
  upcomingMobileSheet.getRange("A3").setValue("No upcoming shows found.").setFontStyle("italic");
  
  SpreadsheetApp.getActiveSpreadsheet().toast("No upcoming or announced shows found.", "Calendar Views Updated");
}

/**
 * Set up headers for all sheets at once
 */
function setUpSheetHeaders(calendarSheet, mobileCalendarSheet, upcomingMobileSheet, totalShowsCount, upcomingShowsCount) {
  // Desktop calendar header
  calendarSheet.getRange("A1").setValue("Concert Calendar").setFontWeight("bold").setFontSize(16);
  
  // Mobile calendar header
  mobileCalendarSheet.getRange("A1").setValue("Mobile Concert Calendar").setFontWeight("bold").setFontSize(14);
  
  // Upcoming shows header with count
  const upcomingHeader = upcomingMobileSheet.getRange(1, 1, 1, 5);
  upcomingHeader.merge()
    .setValue("Upcoming Shows Only")
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("center")
    .setBackground("#20B2AA")
    .setFontColor("white");
  
  // Set count of upcoming shows
  upcomingMobileSheet.getRange(2, 1, 1, 5).merge()
    .setValue(`Total Upcoming Shows: ${upcomingShowsCount}`)
    .setFontWeight("bold")
    .setFontStyle("italic")
    .setHorizontalAlignment("center");
}

/**
 * Process desktop calendar view efficiently
 */
function processDesktopCalendar(sheet, shows, dateInfo, columns) {
  const { monthsInfo } = dateInfo;
  
  // Pre-organize shows by month and day for quick lookup
  const showsByMonthAndDay = {};
  shows.forEach(show => {
    const key = `${show.year}-${show.month}-${show.day}`;
    if (!showsByMonthAndDay[key]) {
      showsByMonthAndDay[key] = [];
    }
    showsByMonthAndDay[key].push(show);
  });
  
  // Process each month
  monthsInfo.forEach((monthInfo, monthIndex) => {
    const monthStartRow = monthIndex * 10 + 3;
    const { year, month, monthYear, firstDay, daysInMonth } = monthInfo;
    
    // Add month header with formatting
    sheet.getRange(monthStartRow, 1, 1, 7).merge()
      .setValue(monthYear)
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground("#E0E0E0")
      .setHorizontalAlignment("center");
    
    // Add day headers with formatting
    sheet.getRange(monthStartRow + 1, 1, 1, 7)
      .setValues([DAYS_OF_WEEK])
      .setFontWeight("bold")
      .setBackground("#F5F5F5")
      .setHorizontalAlignment("center");
    
    // Create the calendar grid efficiently
    const calendarGrid = [];
    let currentWeek = new Array(7).fill("");
    let dayCounter = 1;
    
    // Add empty cells before first day of month
    for (let i = 0; i < firstDay; i++) {
      currentWeek[i] = "";
    }
    
    // Fill in the days
    for (let i = firstDay; i < 7; i++) {
      if (dayCounter <= daysInMonth) {
        currentWeek[i] = dayCounter.toString();
        dayCounter++;
      }
    }
    calendarGrid.push(currentWeek);
    
    // Fill in remaining weeks
    while (dayCounter <= daysInMonth) {
      currentWeek = new Array(7).fill("");
      for (let i = 0; i < 7 && dayCounter <= daysInMonth; i++) {
        currentWeek[i] = dayCounter.toString();
        dayCounter++;
      }
      calendarGrid.push(currentWeek);
    }
    
    // Write calendar grid to sheet and format cells in a batch
    for (let weekIndex = 0; weekIndex < calendarGrid.length; weekIndex++) {
      const weekRange = sheet.getRange(monthStartRow + 2 + weekIndex, 1, 1, 7);
      weekRange.setValues([calendarGrid[weekIndex]]);
      weekRange.setBorder(true, true, true, true, true, true, "#D0D0D0", SpreadsheetApp.BorderStyle.SOLID);
      weekRange.setVerticalAlignment("top");
      weekRange.setHorizontalAlignment("left");
      
      // Set row height for events
      sheet.setRowHeight(monthStartRow + 2 + weekIndex, 80);
    }
    
    // Add show information to cells efficiently
    for (let weekIndex = 0; weekIndex < calendarGrid.length; weekIndex++) {
      for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
        const dayNum = calendarGrid[weekIndex][dayIndex];
        if (dayNum === "") continue;
        
        // Get shows for this day efficiently
        const showKey = `${year}-${month}-${parseInt(dayNum)}`;
        const daysShows = showsByMonthAndDay[showKey] || [];
        
        if (daysShows.length === 0) continue;
        
        // Get cell range
        const cell = sheet.getRange(monthStartRow + 2 + weekIndex, dayIndex + 1);
        
        // Build show info for this cell
        let showInfo = dayNum;
        let cellNote = "";
        
        // Add show information
        daysShows.forEach(show => {
          showInfo += `\n${show.artist}\n${show.venue}`;
          cellNote += `${show.artist} at ${show.venue}\nStatus: ${show.status}\n\n`;
        });
        
        // Update cell with all shows at once
        cell.setValue(showInfo);
        cell.setFontSize(10);
        cell.setBackground(daysShows[0].status === "Upcoming" ? STATUS_COLORS.UPCOMING : STATUS_COLORS.ANNOUNCED);
        cell.setNote(cellNote.trim());
      }
    }
  });
  
  // Add legend for color coding
  const legendRow = monthsInfo.length * 10 + 5;
  sheet.getRange(legendRow, 1).setValue("Color Legend:").setFontWeight("bold");
  sheet.getRange(legendRow, 2).setValue("Upcoming").setBackground(STATUS_COLORS.UPCOMING);
  sheet.getRange(legendRow, 3).setValue("Announced").setBackground(STATUS_COLORS.ANNOUNCED);
}

/**
 * Process mobile calendar view efficiently
 */
function processMobileCalendar(sheet, shows, columns) {
  // Group shows by month for more efficient processing
  const showsByMonth = {};
  shows.forEach(show => {
    const monthYear = `${MONTHS[show.month]} ${show.year}`;
    if (!showsByMonth[monthYear]) {
      showsByMonth[monthYear] = [];
    }
    showsByMonth[monthYear].push(show);
  });
  
  // Sort months chronologically
  const sortedMonths = sortMonthsChronologically(Object.keys(showsByMonth));

  // Process each month in order
  let currentRow = 2;
  sortedMonths.forEach(monthYear => {
    // Add month header
    const monthHeader = sheet.getRange(currentRow, 1, 1, 4);
    monthHeader.merge()
      .setValue(monthYear)
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#E0E0E0")
      .setHorizontalAlignment("center");
    currentRow++;
    
    // Add column headers
    sheet.getRange(currentRow, 1, 1, 4)
      .setValues([["Date", "Artist", "Venue", "Ticket Vendor"]])
      .setFontWeight("bold")
      .setBackground("#F5F5F5")
      .setHorizontalAlignment("center");
    currentRow++;
    
    // Sort shows by date within this month
    const monthShows = showsByMonth[monthYear].sort((a, b) => a.parsedDate - b.parsedDate);
    
    // Prepare data in batches for faster writing
    const showsData = [];
    const backgrounds = [];
    
    monthShows.forEach(show => {
      // Format date for display
      const dayOfWeek = DAYS_OF_WEEK[show.dayOfWeek];
      const formattedDate = `${dayOfWeek} ${show.day}`;
      
      // Get ticket vendor
      const ticketVendor = columns.ticketvendor >= 0 ? show.original[columns.ticketvendor] : "";
      
      // Add to batch arrays
      showsData.push([formattedDate, show.artist, show.venue, ticketVendor]);
      backgrounds.push([show.status === "Upcoming" ? STATUS_COLORS.UPCOMING : STATUS_COLORS.ANNOUNCED]);
    });
    
    // Write and format in batches
    if (showsData.length > 0) {
      sheet.getRange(currentRow, 1, showsData.length, 4).setValues(showsData);
      sheet.getRange(currentRow, 1, showsData.length, 4)
        .setBackgrounds(backgrounds.map(bg => Array(4).fill(bg[0])));
      currentRow += showsData.length;
    }
    
    // Add spacing between months
    currentRow++;
  });
  
  // Add legend
  sheet.getRange(currentRow, 1, 1, 4).merge().setValue("Color Legend:").setFontWeight("bold");
  currentRow++;
  sheet.getRange(currentRow, 1, 1, 4).merge().setValue("Upcoming").setBackground(STATUS_COLORS.UPCOMING);
  currentRow++;
  sheet.getRange(currentRow, 1, 1, 4).merge().setValue("Announced").setBackground(STATUS_COLORS.ANNOUNCED);
}

/**
 * Process upcoming-only mobile view efficiently
 */
function processUpcomingMobileView(sheet, shows, columns) {
  // Check if we have any upcoming shows
  if (shows.length === 0) {
    sheet.getRange(3, 1).setValue("No upcoming shows found.").setFontStyle("italic");
    return;
  }
  
  // Group shows by month for efficient processing
  const showsByMonth = {};
  shows.forEach(show => {
    const monthYear = `${MONTHS[show.month]} ${show.year}`;
    if (!showsByMonth[monthYear]) {
      showsByMonth[monthYear] = [];
    }
    showsByMonth[monthYear].push(show);
  });
  
  // Sort months chronologically
  const sortedMonths = sortMonthsChronologically(Object.keys(showsByMonth));

  // Get current date for days calculation
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Process each month
  let currentRow = 3;
  sortedMonths.forEach(monthYear => {
    // Add month header with better mobile formatting
    const monthHeader = sheet.getRange(currentRow, 1, 1, 5);
    monthHeader.merge()
      .setValue(monthYear)
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#4682B4") // Steel Blue for more distinct month header
      .setFontColor("white")
      .setHorizontalAlignment("center");
    currentRow++;
    
    // Add column headers
    sheet.getRange(currentRow, 1, 1, 5)
      .setValues([["Date", "Days", "Artist", "Venue", "Ticket"]])
      .setFontWeight("bold")
      .setBackground("#F5F5F5")
      .setHorizontalAlignment("center");
    currentRow++;
    
    // Sort shows by date within this month
    const monthShows = showsByMonth[monthYear].sort((a, b) => a.parsedDate - b.parsedDate);
    
    // Prepare batch data for faster writing
    const showsData = [];
    
    // Track which rows need special formatting
    const todayRows = [];
    const soonRows = [];
    
    monthShows.forEach(show => {
      // Format date as "Mon 13"
      const dayOfWeek = DAYS_OF_WEEK[show.dayOfWeek].substring(0, 3);
      const formattedDate = `${dayOfWeek} ${show.day}`;
      
      // Calculate days until show
      const concertDate = new Date(show.parsedDate);
      concertDate.setHours(0, 0, 0, 0);
      const timeDiff = concertDate.getTime() - today.getTime();
      const diffDays = Math.ceil(timeDiff / (1000 * 60 * 60 * 24));
      
      // Create concise days counter text
      let daysUntilText = diffDays === 0 ? "TODAY!" : 
                        diffDays === 1 ? "TMR!" : 
                        diffDays.toString();
      
      // Get ticket vendor
      const ticketVendor = columns.ticketvendor >= 0 ? show.original[columns.ticketvendor] : "";
      
      // Add to data array
      showsData.push([formattedDate, daysUntilText, show.artist, show.venue, ticketVendor]);
      
      // Track rows needing special formatting
      if (diffDays === 0) {
        todayRows.push(showsData.length - 1);
      } else if (diffDays <= 3) {
        soonRows.push(showsData.length - 1);
      }
    });
    
    // Write all values at once
    if (showsData.length > 0) {
      const dataRange = sheet.getRange(currentRow, 1, showsData.length, 5);
      dataRange.setValues(showsData);
      dataRange.setBackgrounds(
        showsData.map((_, i) => Array(5).fill(i % 2 === 0 ? "#E6F3F2" : "#CEEAE8"))
      );

      // Apply special formatting for today/soon shows
      todayRows.forEach(rowIndex => {
        sheet.getRange(currentRow + rowIndex, 2).setFontWeight("bold").setFontColor("#FF0000");
      });
      soonRows.forEach(rowIndex => {
        sheet.getRange(currentRow + rowIndex, 2).setFontWeight("bold");
      });

      currentRow += showsData.length;
    }
    
    // Add spacing between months
    currentRow++;
  });
  
  // Add helpful notes
  currentRow++;
  sheet.getRange(currentRow, 1, 1, 5).merge()
    .setValue("Days column shows number of days until the concert")
    .setFontStyle("italic")
    .setFontSize(10)
    .setHorizontalAlignment("center");
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 5).merge()
    .setValue("Alternating row colors for easier reading on mobile devices")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center");
  currentRow++;
  
  // Add timestamp
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy h:mm a");
  sheet.getRange(currentRow + 1, 1, 1, 5).merge()
    .setValue(`Last updated: ${timestamp}`)
    .setFontStyle("italic")
    .setHorizontalAlignment("center");
}

/**
 * Format all sheets in batch operations
 */
function formatSheets(calendarSheet, mobileCalendarSheet, upcomingMobileSheet) {
  // Format desktop calendar
  calendarSheet.setColumnWidths(1, 7, 150);
  
  // Format mobile calendar
  mobileCalendarSheet.setColumnWidth(1, 100);  // Date column
  mobileCalendarSheet.setColumnWidth(2, 150);  // Artist column
  mobileCalendarSheet.setColumnWidth(3, 150);  // Venue column
  mobileCalendarSheet.setColumnWidth(4, 150);  // Ticket Vendor column
  
  // Format upcoming mobile view
  upcomingMobileSheet.setColumnWidth(1, 55);   // Date column - compact
  upcomingMobileSheet.setColumnWidth(2, 45);   // Days Until column - compact
  upcomingMobileSheet.setColumnWidth(3, 120);  // Artist column
  upcomingMobileSheet.setColumnWidth(4, 120);  // Venue column
  upcomingMobileSheet.setColumnWidth(5, 90);   // Ticket vendor column - compact
}

/**
 * Helper function to set up a section header in the analysis sheet
 */
function setupAnalysisSection(sheet, row, title, fontSize) {
  sheet.getRange(row, 1).setValue(title).setFontWeight("bold").setFontSize(fontSize);
}

/**
 * Add an analysis table to the sheet
 */
function addAnalysisTable(sheet, startRow, title, dataRows) {
  // Safety check for dataRows
  if (!dataRows || !Array.isArray(dataRows) || dataRows.length === 0) {
    console.error("Invalid dataRows provided to addAnalysisTable");
    sheet.getRange(startRow, 1).setValue(title).setFontWeight("bold").setFontSize(12);
    sheet.getRange(startRow + 1, 1).setValue("No data available");
    return 2; // Return minimal rows used
  }
  
  setupAnalysisSection(sheet, startRow, title, 12);
  
  // Format table header with blue background and white text
  sheet.getRange(startRow + 1, 1, 1, dataRows[0].length).setValues([dataRows[0]])
    .setFontWeight("bold")
    .setBackground(ANALYSIS_COLORS.HEADER_BG)
    .setFontColor(ANALYSIS_COLORS.HEADER_TEXT)
    .setHorizontalAlignment("center");
  
  // Add data rows
  if (dataRows.length > 1) {
    sheet.getRange(startRow + 2, 1, dataRows.length - 1, dataRows[0].length)
      .setValues(dataRows.slice(1));
  }
  
  return dataRows.length;
}