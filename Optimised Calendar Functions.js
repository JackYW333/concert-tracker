/**
 * Safely shows a toast message if called from a UI context
 * @param {string} message - The message to display
 * @param {string} title - The toast title
 * @param {number} timeoutSeconds - How long to show the toast (optional)
 */
function safeToast(message, title, timeoutSeconds = 5) {
  try {
    // Only attempt to show toast if we're in a UI context
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeoutSeconds);
  } catch (e) {
    // If we can't show a toast (e.g., in a time-driven trigger), just log it
    Logger.log(`${title}: ${message}`);
  }
}
/**
 * Optimized concert calendar generation functions
 * Separates calendar views into individual functions to prevent timeouts
 */

// Use existing constants if defined, otherwise create new ones
// This avoids "variable already declared" errors
const CALENDAR_CONSTANTS = (function() {
  // Check if constants already exist in the global scope
  const existingStatusColors = typeof STATUS_COLORS !== 'undefined' ? STATUS_COLORS : {
    UPCOMING: "#20B2AA", // Teal
    ANNOUNCED: "#FFA500" // Orange
  };
  
  const existingSheetNames = typeof SHEET_NAMES !== 'undefined' ? SHEET_NAMES : {
    MAIN: "Tracker",
    CALENDAR: "Calendar View",
    MOBILE_CALENDAR: "Mobile Calendar",
    UPCOMING_MOBILE: "Upcoming Mobile Calendar"
  };
  
  // Use existing constants or create new ones for days and months
  const existingDaysOfWeek = typeof DAYS_OF_WEEK !== 'undefined' ? DAYS_OF_WEEK : 
    ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    
  const existingMonths = typeof MONTHS !== 'undefined' ? MONTHS : 
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  
  return {
    STATUS_COLORS: existingStatusColors,
    SHEET_NAMES: existingSheetNames,
    DAYS_OF_WEEK: existingDaysOfWeek,
    MONTHS: existingMonths
  };
})();

/**
 * Helper function to get or create a sheet
 */
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * Helper function to parse a date value (handles string dates and Date objects)
 * @param {any} dateValue - The date value (string or Date)
 * @return {Date|null} - Parsed Date object or null if invalid
 */
function parseDate(dateValue) {
  // Already a Date object
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
    return dateValue;
  }
  
  // String date in dd/mm/yyyy format
  if (typeof dateValue === 'string' && dateValue.includes('/')) {
    try {
      const dateParts = dateValue.split('/');
      if (dateParts.length === 3) {
        const day = parseInt(dateParts[0], 10);
        const month = parseInt(dateParts[1], 10) - 1;
        const year = parseInt(dateParts[2], 10);
        
        const parsedDate = new Date(year, month, day);
        if (!isNaN(parsedDate.getTime())) {
          return parsedDate;
        }
      }
    } catch (e) {
      console.error("Error parsing date:", e);
    }
  }
  
  // Try standard date parsing
  try {
    const parsedDate = new Date(dateValue);
    if (!isNaN(parsedDate.getTime())) {
      return parsedDate;
    }
  } catch (e) {
    // Invalid date format
  }
  
  // Could not parse date
  return null;
}

/**
 * Helper function to get all upcoming/announced shows with parsed dates
 * @param {boolean} upcomingOnly - If true, only returns shows with "Upcoming" status
 * @return {Object} - Object containing shows and column indexes
 */
function getShowData(upcomingOnly = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CALENDAR_CONSTANTS.SHEET_NAMES.MAIN);
  
  if (!mainSheet) {
    throw new Error("Main tracker sheet not found");
  }
  
  // Get data in one batch operation
  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indexes
  const columns = {
    artist: headers.indexOf("Artist"),
    date: headers.indexOf("Date"),
    venue: headers.indexOf("Venue"),
    city: headers.indexOf("City"),
    status: headers.indexOf("Status"),
    ticketVendor: headers.indexOf("Ticket Vendor")
  };
  
  // Validate required columns
  if (columns.artist === -1 || columns.date === -1 || columns.status === -1) {
    throw new Error("Required columns missing (Artist, Date, or Status)");
  }
  
  // Filter for upcoming/announced shows
  const statusFilter = upcomingOnly ? 
    (status) => status === "Upcoming" : 
    (status) => status === "Upcoming" || status === "Announced";
  
  // Process shows
  const shows = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[columns.status];
    
    // Check status
    if (!statusFilter(status)) continue;
    
    // Parse date
    const parsedDate = parseDate(row[columns.date]);
    if (!parsedDate) continue;
    
    // Add show with parsed date
    shows.push({
      original: row,
      artist: row[columns.artist],
      venue: row[columns.venue],
      city: columns.city !== -1 ? row[columns.city] : "",
      status: status,
      ticketVendor: columns.ticketVendor !== -1 ? row[columns.ticketVendor] : "",
      parsedDate: parsedDate,
      month: parsedDate.getMonth(),
      year: parsedDate.getFullYear(),
      day: parsedDate.getDate(),
      dayOfWeek: parsedDate.getDay()
    });
  }
  
  // Sort shows by date
  shows.sort((a, b) => a.parsedDate - b.parsedDate);
  
  return { shows, columns };
}

/**
 * Generate desktop calendar view with month-by-month grid
 */
function generateDesktopCalendar() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const calendarSheet = getOrCreateSheet(CALENDAR_CONSTANTS.SHEET_NAMES.CALENDAR);
    
    // Clear existing data
    calendarSheet.clear();
    
    // Set title
    calendarSheet.getRange("A1").setValue("Concert Calendar")
      .setFontWeight("bold")
      .setFontSize(16);
    
    // Get shows data
    const { shows } = getShowData();
    
    if (shows.length === 0) {
      calendarSheet.getRange("A3").setValue("No upcoming or announced shows found.")
        .setFontStyle("italic");
      safeToast("No upcoming or announced shows found.", "Calendar Updated");
      return;
    }
    
    // Get current date and max date
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Find max date
    let maxDate = new Date(today);
    shows.forEach(show => {
      if (show.parsedDate > maxDate) {
        maxDate = new Date(show.parsedDate);
      }
    });
    
    // Calculate months to display
    const currentMonth = today.getMonth() + today.getFullYear() * 12;
    const maxMonth = maxDate.getMonth() + maxDate.getFullYear() * 12;
    const monthsToDisplay = Math.max(1, maxMonth - currentMonth + 1);
    
    // Pre-organize shows by month-day key for quick lookup
    const showsByMonthDay = {};
    shows.forEach(show => {
      const key = `${show.year}-${show.month}-${show.day}`;
      if (!showsByMonthDay[key]) {
        showsByMonthDay[key] = [];
      }
      showsByMonthDay[key].push(show);
    });
    
    // Generate months in batches
    for (let i = 0; i < monthsToDisplay; i++) {
      const monthDate = new Date(today);
      monthDate.setMonth(today.getMonth() + i);
      
      const year = monthDate.getFullYear();
      const month = monthDate.getMonth();
      const monthName = CALENDAR_CONSTANTS.MONTHS[month];
      const monthYear = `${monthName} ${year}`;
      
      const firstDayOfMonth = new Date(year, month, 1).getDay();
      const daysInMonth = new Date(year, month + 1, 0).getDate();
      
      const monthStartRow = i * 10 + 3;
      
      // Add month header
      const headerRange = calendarSheet.getRange(monthStartRow, 1, 1, 7);
      headerRange.merge()
        .setValue(monthYear)
        .setFontWeight("bold")
        .setFontSize(14)
        .setBackground("#E0E0E0")
        .setHorizontalAlignment("center");
      
      // Add day headers
      calendarSheet.getRange(monthStartRow + 1, 1, 1, 7)
        .setValues([CALENDAR_CONSTANTS.DAYS_OF_WEEK])
        .setFontWeight("bold")
        .setBackground("#F5F5F5")
        .setHorizontalAlignment("center");
      
      // Create calendar grid efficiently
      const calendarGrid = [];
      let dayCounter = 1;
      
      // Week 1 with empty cells for days before month starts
      let week1 = new Array(7).fill("");
      for (let d = firstDayOfMonth; d < 7 && dayCounter <= daysInMonth; d++) {
        week1[d] = dayCounter.toString();
        dayCounter++;
      }
      calendarGrid.push(week1);
      
      // Remaining weeks
      while (dayCounter <= daysInMonth) {
        let week = new Array(7).fill("");
        for (let d = 0; d < 7 && dayCounter <= daysInMonth; d++) {
          week[d] = dayCounter.toString();
          dayCounter++;
        }
        calendarGrid.push(week);
      }
      
      // Add calendar grid to sheet
      for (let weekIdx = 0; weekIdx < calendarGrid.length; weekIdx++) {
        const row = monthStartRow + 2 + weekIdx;
        
        // Set values for the entire week at once
        calendarSheet.getRange(row, 1, 1, 7)
          .setValues([calendarGrid[weekIdx]])
          .setBorder(true, true, true, true, true, true, "#D0D0D0", SpreadsheetApp.BorderStyle.SOLID)
          .setVerticalAlignment("top")
          .setHorizontalAlignment("left");
        
        // Set row height for events
        calendarSheet.setRowHeight(row, 80);
      }
      
      // Add shows to calendar
      for (let weekIdx = 0; weekIdx < calendarGrid.length; weekIdx++) {
        for (let dayIdx = 0; dayIdx < 7; dayIdx++) {
          const dayNum = calendarGrid[weekIdx][dayIdx];
          if (dayNum === "") continue;
          
          // Get shows for this day
          const dayKey = `${year}-${month}-${parseInt(dayNum)}`;
          const dayShows = showsByMonthDay[dayKey] || [];
          
          if (dayShows.length === 0) continue;
          
          // Update cell with show info
          const cell = calendarSheet.getRange(monthStartRow + 2 + weekIdx, dayIdx + 1);
          
          let showInfo = dayNum;
          let cellNote = "";
          
          dayShows.forEach(show => {
            showInfo += `\n${show.artist}\n${show.venue}`;
            cellNote += `${show.artist} at ${show.venue}\nStatus: ${show.status}\n\n`;
          });
          
          cell.setValue(showInfo)
            .setFontSize(10)
            .setBackground(dayShows[0].status === "Upcoming" ? CALENDAR_CONSTANTS.STATUS_COLORS.UPCOMING : CALENDAR_CONSTANTS.STATUS_COLORS.ANNOUNCED)
            .setNote(cellNote.trim());
        }
      }
    }
    
    // Set column widths
    calendarSheet.setColumnWidths(1, 7, 150);
    
    // Add legend
    const legendRow = monthsToDisplay * 10 + 5;
    calendarSheet.getRange(legendRow, 1).setValue("Color Legend:").setFontWeight("bold");
    calendarSheet.getRange(legendRow, 2).setValue("Upcoming").setBackground(CALENDAR_CONSTANTS.STATUS_COLORS.UPCOMING);
    calendarSheet.getRange(legendRow, 3).setValue("Announced").setBackground(CALENDAR_CONSTANTS.STATUS_COLORS.ANNOUNCED);
    
    safeToast("Desktop calendar view updated successfully!", "Calendar Updated");
  } catch (error) {
    console.error("Error generating desktop calendar:", error);
    safeToast("Error generating desktop calendar: " + error.message, "Error");
  }
}

/**
 * Generate mobile-friendly calendar with month-by-month list
 */
function generateMobileCalendar() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mobileSheet = getOrCreateSheet(CALENDAR_CONSTANTS.SHEET_NAMES.MOBILE_CALENDAR);
    
    // Clear existing data
    mobileSheet.clear();
    
    // Set title
    mobileSheet.getRange("A1").setValue("Mobile Concert Calendar")
      .setFontWeight("bold")
      .setFontSize(14);
    
    // Get shows data
    const { shows } = getShowData();
    
    if (shows.length === 0) {
      mobileSheet.getRange("A3").setValue("No upcoming or announced shows found.")
        .setFontStyle("italic");
      safeToast("No upcoming or announced shows found.", "Mobile Calendar Updated");
      return;
    }
    
    // Group shows by month for more efficient processing
    const showsByMonth = {};
    shows.forEach(show => {
      const monthYear = `${CALENDAR_CONSTANTS.MONTHS[show.month]} ${show.year}`;
      if (!showsByMonth[monthYear]) {
        showsByMonth[monthYear] = [];
      }
      showsByMonth[monthYear].push(show);
    });
    
    // Sort months chronologically
    const sortedMonths = sortMonthsChronologically(Object.keys(showsByMonth));

    // Process months
    let currentRow = 2;
    sortedMonths.forEach(monthYear => {
      // Add month header
      const headerRange = mobileSheet.getRange(currentRow, 1, 1, 4);
      headerRange.merge()
        .setValue(monthYear)
        .setFontWeight("bold")
        .setFontSize(12)
        .setBackground("#E0E0E0")
        .setHorizontalAlignment("center");
      currentRow++;
      
      // Add column headers
      mobileSheet.getRange(currentRow, 1, 1, 4)
        .setValues([["Date", "Artist", "Venue", "Ticket Vendor"]])
        .setFontWeight("bold")
        .setBackground("#F5F5F5")
        .setHorizontalAlignment("center");
      currentRow++;
      
      // Sort shows by date within this month
      const monthShows = showsByMonth[monthYear].sort((a, b) => a.parsedDate - b.parsedDate);
      
      // Prepare batch data for efficient writing
      const showsData = [];
      const backgrounds = [];
      
      monthShows.forEach(show => {
        // Format date for display
        const dayOfWeek = CALENDAR_CONSTANTS.DAYS_OF_WEEK[show.dayOfWeek];
        const formattedDate = `${dayOfWeek} ${show.day}`;
        
        // Add to batch arrays
        showsData.push([formattedDate, show.artist, show.venue, show.ticketVendor]);
        backgrounds.push(show.status === "Upcoming" ? CALENDAR_CONSTANTS.STATUS_COLORS.UPCOMING : CALENDAR_CONSTANTS.STATUS_COLORS.ANNOUNCED);
      });
      
      // Write data in batches for better performance
      if (showsData.length > 0) {
        const dataRange = mobileSheet.getRange(currentRow, 1, showsData.length, 4);
        dataRange.setValues(showsData);
        dataRange.setBackgrounds(backgrounds.map(bg => Array(4).fill(bg)));
        currentRow += showsData.length;
      }
      
      // Add spacing between months
      currentRow++;
    });
    
    // Add legend
    mobileSheet.getRange(currentRow, 1, 1, 4).merge().setValue("Color Legend:").setFontWeight("bold");
    currentRow++;
    mobileSheet.getRange(currentRow, 1, 1, 4).merge().setValue("Upcoming").setBackground(CALENDAR_CONSTANTS.STATUS_COLORS.UPCOMING);
    currentRow++;
    mobileSheet.getRange(currentRow, 1, 1, 4).merge().setValue("Announced").setBackground(CALENDAR_CONSTANTS.STATUS_COLORS.ANNOUNCED);
    
    // Format mobile calendar
    mobileSheet.setColumnWidth(1, 100); // Date column
    mobileSheet.setColumnWidth(2, 150); // Artist column
    mobileSheet.setColumnWidth(3, 150); // Venue column
    mobileSheet.setColumnWidth(4, 150); // Ticket Vendor column
    
    safeToast("Mobile calendar view updated successfully!", "Mobile Calendar Updated");
  } catch (error) {
    console.error("Error generating mobile calendar:", error);
    safeToast("Error generating mobile calendar: " + error.message, "Error");
  }
}

/**
 * Generate upcoming-only mobile view with dark mode compatible colors
 */
function generateUpcomingMobileView() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Make sure we're using the correct sheet name "Upcoming Mobile Calendar"
    const upcomingSheetName = "Upcoming Mobile Calendar";
    const upcomingSheet = getOrCreateSheet(upcomingSheetName);
    
    // Clear existing data
    upcomingSheet.clear();
    
    // Set title with dark-mode friendly formatting - using deeper teal
    upcomingSheet.getRange(1, 1, 1, 5).merge()
      .setValue("Upcoming Shows Only")
      .setFontWeight("bold")
      .setFontSize(16)
      .setHorizontalAlignment("center")
      .setBackground("#008080") // Darker teal color for better dark mode contrast
      .setFontColor("white");
    
    // Get upcoming-only shows
    const { shows } = getShowData(true);
    
    // Display count of upcoming shows
    upcomingSheet.getRange(2, 1, 1, 5).merge()
      .setValue(`Total Upcoming Shows: ${shows.length}`)
      .setFontWeight("bold")
      .setFontStyle("italic")
      .setHorizontalAlignment("center")
      .setBackground("#263238")
      .setFontColor("white");
    
    if (shows.length === 0) {
      upcomingSheet.getRange(3, 1).setValue("No upcoming shows found.")
        .setFontStyle("italic");
      safeToast("No upcoming shows found.", "Upcoming Shows Updated");
      return;
    }
    
    // Group shows by month
    const showsByMonth = {};
    shows.forEach(show => {
      const monthYear = `${CALENDAR_CONSTANTS.MONTHS[show.month]} ${show.year}`;
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
      // Add month header with dark-mode friendly formatting
      const headerRange = upcomingSheet.getRange(currentRow, 1, 1, 5);
      headerRange.merge()
        .setValue(monthYear)
        .setFontWeight("bold")
        .setFontSize(12)
        .setBackground("#0D6E6E") // Darker teal for month headers - better for dark mode
        .setFontColor("white")
        .setHorizontalAlignment("center");
      currentRow++;
      
      // Add column headers
      const colHeaders = upcomingSheet.getRange(currentRow, 1, 1, 5);
      colHeaders.setValues([["Date", "Days", "Artist", "Venue", "Ticket"]])
        .setFontWeight("bold")
        .setBackground("#263238") // Dark blue-grey background for headers
        .setFontColor("white")    // White text for contrast
        .setHorizontalAlignment("center");
      currentRow++;
      
      // Sort shows by date within this month
      const monthShows = showsByMonth[monthYear].sort((a, b) => a.parsedDate - b.parsedDate);
      
      // Prepare batch data
      const showsData = [];
      const todayRows = [];
      const soonRows = [];
      
      monthShows.forEach(show => {
        // Format date - short day name + day number
        const dayOfWeek = CALENDAR_CONSTANTS.DAYS_OF_WEEK[show.dayOfWeek].substring(0, 3);
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
        
        showsData.push({
          values: [formattedDate, daysUntilText, show.artist, show.venue, show.ticketVendor],
          diffDays: diffDays
        });
        
        // Track rows needing special formatting
        if (diffDays === 0) {
          todayRows.push(showsData.length - 1);
        } else if (diffDays <= 3) {
          soonRows.push(showsData.length - 1);
        }
      });
      
      // Write data in batches for better performance
      if (showsData.length > 0) {
        const values = showsData.map(item => item.values);
        const dataRange = upcomingSheet.getRange(currentRow, 1, values.length, 5);
        dataRange.setValues(values);
        dataRange.setBackgrounds(
          values.map((_, i) => Array(5).fill(i % 2 === 0 ? "#195959" : "#0A4040"))
        );
        dataRange.setFontColors(values.map(() => Array(5).fill("white")));

        // Special formatting for today/soon shows
        todayRows.forEach(i => {
          upcomingSheet.getRange(currentRow + i, 2)
            .setFontWeight("bold")
            .setFontColor("#FF9E80");
        });
        soonRows.forEach(i => {
          upcomingSheet.getRange(currentRow + i, 2)
            .setFontWeight("bold")
            .setFontColor("#80DEEA");
        });

        currentRow += showsData.length;
      }
      
      // Add spacing between months with a dark divider color
      const dividerRange = upcomingSheet.getRange(currentRow, 1, 1, 5);
      dividerRange.merge()
        .setBackground("#0A2A2A"); // Very dark teal as a divider
      upcomingSheet.setRowHeight(currentRow, 4); // Make divider thin but visible
      currentRow++;
    });
    
    // Add helpful notes with dark mode friendly styling
    currentRow++;
    upcomingSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue("Days column shows number of days until the concert")
      .setFontStyle("italic")
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setFontColor("#B0BEC5"); // Light grey-blue for notes
    currentRow++;
    
    upcomingSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue("Dark mode compatible colors for better readability")
      .setFontSize(10)
      .setFontStyle("italic")
      .setHorizontalAlignment("center")
      .setFontColor("#B0BEC5"); // Light grey-blue for notes
    currentRow++;
    
    // Add timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy h:mm a");
    upcomingSheet.getRange(currentRow + 1, 1, 1, 5).merge()
      .setValue(`Last updated: ${timestamp}`)
      .setFontStyle("italic")
      .setHorizontalAlignment("center")
      .setFontColor("#B0BEC5"); // Light grey-blue for timestamp
    
    // Format columns for better mobile view
    upcomingSheet.setColumnWidth(1, 55);   // Date column - compact
    upcomingSheet.setColumnWidth(2, 45);   // Days Until column - compact
    upcomingSheet.setColumnWidth(3, 120);  // Artist column
    upcomingSheet.setColumnWidth(4, 120);  // Venue column
    upcomingSheet.setColumnWidth(5, 90);   // Ticket vendor column - compact
    
    safeToast("Upcoming shows view updated successfully!", "Upcoming Shows Updated");
  } catch (error) {
    console.error("Error generating upcoming shows view:", error);
    safeToast("Error generating upcoming shows view: " + error.message, "Error");
  }
}

/**
 * Generate all calendar views with progress tracking
 */
function generateAllCalendarViews() {
  try {
    // Step 1: Generate desktop calendar
    safeToast("Generating desktop calendar view...", "Progress 1/3");
    generateDesktopCalendar();
    
    // Step 2: Generate mobile calendar
    safeToast("Generating mobile calendar view...", "Progress 2/3");
    generateMobileCalendar();
    
    // Step 3: Generate upcoming shows view
    safeToast("Generating upcoming shows view...", "Progress 3/3");
    generateUpcomingMobileView();
    
    // Show completion message
    safeToast("All calendar views updated successfully!", "Complete");
  } catch (error) {
    console.error("Error generating calendar views:", error);
    safeToast("Error generating calendar views: " + error.message, "Error");
  }
}