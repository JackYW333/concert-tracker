/**
 * Script to transfer form response data to the Tracker sheet
 * Creates rows for each Artist with appropriate data mapping
 */
function transferFormResponsesToTracker(e) {
  // Check if this is running from a trigger
  const isTriggerDriven = (e !== undefined);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the form responses sheet - check multiple possible names
  let formSheet = ss.getSheetByName("Form Responses 1");
  if (!formSheet) {
    formSheet = ss.getSheetByName("Form Responses");
    if (!formSheet) {
      // Try to find any sheet with "Form" and "Response" in the name
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const sheetName = sheets[i].getName();
        if (sheetName.toLowerCase().includes("form") && 
            sheetName.toLowerCase().includes("response")) {
          formSheet = sheets[i];
          break;
        }
      }
    }
  }
  
  // Get the tracker sheet
  const trackerSheet = ss.getSheetByName("Tracker");
  
  // Check if sheets were found
  if (!formSheet) {
    Logger.log("Form response sheet not found.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: Form response sheet not found.", "Error", 5);
    return;
  }
  
  if (!trackerSheet) {
    Logger.log("Tracker sheet not found.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: Tracker sheet not found.", "Error", 5);
    return;
  }
  
  // If trigger-driven, use the trigger event data
  if (isTriggerDriven) {
    try {
      processFormResponse(e.values, formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0], trackerSheet);
    } catch (error) {
      Logger.log("Error processing form submission: " + error.toString());
      SpreadsheetApp.getActiveSpreadsheet().toast("Error processing form submission. Check logs for details.", "Error", 5);
    }
  } else {
    // Otherwise process the latest response manually
    try {
      processLatestFormResponse(formSheet, trackerSheet);
    } catch (error) {
      Logger.log("Error processing latest form response: " + error.toString());
      SpreadsheetApp.getActiveSpreadsheet().toast("Error processing latest form response. Check logs for details.", "Error", 5);
    }
  }
}

/**
 * Process the latest form response and transfer data to Tracker
 */
function processLatestFormResponse(formSheet, trackerSheet) {
  // Check if form sheet exists
  if (!formSheet) {
    Logger.log("Form sheet not found. Make sure the sheet exists and is named correctly.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: Form response sheet not found.", "Error", 5);
    return;
  }
  
  const lastRow = formSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("No form responses found (only header row exists).");
    return; // Only header row exists
  }
  
  // Get headers and the last form submission
  const headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
  const formData = formSheet.getRange(lastRow, 1, 1, formSheet.getLastColumn()).getValues()[0];
  
  // Process this response
  processFormResponse(formData, headers, trackerSheet);
}

/**
 * Process a single form response
 */
function processFormResponse(formData, headers, trackerSheet) {
  // Create an object with column names as keys
  const response = {};
  headers.forEach((header, index) => {
    response[header] = formData[index];
  });
  
  // Define artist columns (using column indices - K, M, O, etc.)
  const artistColumns = [
    {name: 10, rating: 11}, // K, L (0-indexed)
    {name: 12, rating: 13}, // M, N
    {name: 14, rating: 15}, // O, P
    {name: 16, rating: 17}, // Q, R
    {name: 18, rating: 19}, // S, T
    // U is Position, not an artist column
    {name: 21, rating: 22}, // V, W
    {name: 23, rating: 24}, // X, Y
    {name: 25, rating: 26}, // Z, AA
    {name: 27, rating: 28}, // AB, AC
    {name: 29, rating: 30}, // AD, AE
    {name: 31, rating: 32}, // AF, AG
    {name: 33, rating: 34}, // AH, AI
    {name: 35, rating: 36}, // AJ, AK
    {name: 37, rating: 38}, // AL, AM
    {name: 39, rating: 40}, // AN, AO
    {name: 41, rating: null}  // AP (no rating column)
  ];
  
  // Get the type of show to determine roles
  const typeOfShow = response[headers[7]]; // Column H (8th column, 0-indexed is 7)
  
  // For Concert shows, the headline artist is from Column K (index 10)
  let headlineArtist = null;
  if (typeOfShow === "Concert") {
    headlineArtist = formData[10]; // Column K - the headline artist
  }
  
  // Get position from column U
  const position = formData[20]; // Column U (21st column, 0-indexed is 20)
  
  // Original event name
  const eventName = formData[8]; // Column I (0-indexed is 8)
  
  // Get the original date from the form (Column B)
  let dateValue = formData[1]; // Column B (0-indexed is 1)
  
  // Make sure we have a proper Date object
  // This is critical for the calendar view to work properly
  if (!(dateValue instanceof Date)) {
    if (typeof dateValue === 'string') {
      try {
        // Try to parse the date string (assuming mm/dd/yyyy format from the form)
        const dateParts = dateValue.split('/');
        if (dateParts.length === 3) {
          const month = parseInt(dateParts[0], 10);
          const day = parseInt(dateParts[1], 10);
          const year = parseInt(dateParts[2], 10);
          
          // Create a proper Date object
          dateValue = new Date(year, month - 1, day);
        }
      } catch (e) {
        console.error("Error parsing date string: ", e);
      }
    }
  }
  
  // Extract all artists that have data
  const artists = [];
  artistColumns.forEach((cols, index) => {
    const artistName = formData[cols.name];
    if (artistName && artistName.toString().trim() !== "") {
      const rating = cols.rating !== null ? formData[cols.rating] : "";
      artists.push({
        name: artistName,
        rating: rating,
        isColumnK: cols.name === 10 // Check if this is the artist from column K
      });
    }
  });
  
  // Create rows in tracker for each artist
  artists.forEach(artist => {
    // Determine the role based on show type and column
    let role = "";
    if (artist.isColumnK) {
      // Artist from Column K should always be "Headliner"
      role = "Headliner";
    } else if (typeOfShow === "Festival") {
      role = "Festival Act";
    } else if (typeOfShow === "Concert") {
      role = (artist.name === headlineArtist) ? "Headline" : "Support";
    }
    
    // Determine event name - for support artists, add headline artist
    let modifiedEventName = eventName;
    if (role === "Support" && headlineArtist) {
      modifiedEventName = `${eventName} ${headlineArtist}`;
    }
    
    // Create row data for tracker based on the requested column order:
    // Artist, Role, Date, Event Name, Venue, City, Who I Went With, Notes, Position, Rating, Status, Ticket Vendor, Ticket Price
    const newRow = [
      artist.name,                       // Artist
      role,                              // Role (now "Headliner" for Column K artists)
      dateValue,                         // Date (Column B) - keeping as a Date object for proper calendar functionality
      modifiedEventName,                 // Event Name (Column I) + headline artist for support acts
      formData[2],                       // Venue (Column C)
      formData[3],                       // City (Column D)
      formData[6],                       // Who I Went With (Column G)
      formData[9],                       // Notes (Column J)
      position,                          // Position (Column U)
      artist.rating,                     // Rating
      formData[5],                       // Status (Column F)
      formData[formData.length - 1],     // Ticket Vendor (Column AQ)
      formData[4]                        // Ticket Price (Column E)
    ];
    
    // Add the row to the tracker sheet
    trackerSheet.appendRow(newRow);
  });
}

/**
 * Clean up existing triggers before creating a new one
 * This helps avoid duplicate triggers
 */
function cleanupExistingTriggers() {
  // Delete all existing triggers to avoid duplicates
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'transferFormResponsesToTracker') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
}

/**
 * Creates a trigger to run this script when a form is submitted
 */
function createFormSubmitTrigger() {
  // Clean up existing triggers first
  cleanupExistingTriggers();
  
  // Create a new trigger
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('transferFormResponsesToTracker')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  
  Logger.log("Form submit trigger created successfully");
}

/**
 * Process all existing form responses
 * Use this function to process historical data
 */
function processAllExistingResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form Responses 1");
  const trackerSheet = ss.getSheetByName("Tracker");
  
  // Check if sheets exist
  if (!formSheet) {
    Logger.log("Form Responses 1 sheet not found");
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: Form Responses 1 sheet not found", "Error");
    return;
  }
  
  if (!trackerSheet) {
    Logger.log("Tracker sheet not found");
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: Tracker sheet not found", "Error");
    return;
  }
  
  // Get all data from the form responses sheet
  const data = formSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("No data found (only header row exists)");
    SpreadsheetApp.getActiveSpreadsheet().toast("No data to process", "Info");
    return; // Only header row exists
  }
  
  const headers = data[0];
  
  // Process each row, starting from row 2 (after headers)
  try {
    let processedCount = 0;
    for (let i = 1; i < data.length; i++) {
      processFormResponse(data[i], headers, trackerSheet);
      processedCount++;
      Utilities.sleep(500); // Add a small delay to prevent rate limiting
    }
    SpreadsheetApp.getActiveSpreadsheet().toast(`Successfully processed ${processedCount} form responses`, "Success");
  } catch (error) {
    Logger.log("Error processing form responses: " + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast("Error processing form responses: " + error.toString(), "Error");
  }
}

/**
 * This function can be run directly to process the latest form response
 */
function processLatestFormResponseDirect() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form Responses 1");
  const trackerSheet = ss.getSheetByName("Tracker");
  
  if (!formSheet) {
    Logger.log("Error: Form Responses 1 sheet not found");
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: Form Responses 1 sheet not found", "Error");
    return;
  }
  
  if (!trackerSheet) {
    Logger.log("Error: Tracker sheet not found");
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: Tracker sheet not found", "Error");
    return;
  }
  
  const lastRow = formSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("No data in Form Responses 1 (only header row found)");
    SpreadsheetApp.getActiveSpreadsheet().toast("No data found in Form Responses 1", "Info");
    return;
  }
  
  // Get headers and the last form submission
  const headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
  const formData = formSheet.getRange(lastRow, 1, 1, formSheet.getLastColumn()).getValues()[0];
  
  try {
    // Process this response
    processFormResponse(formData, headers, trackerSheet);
    SpreadsheetApp.getActiveSpreadsheet().toast("Successfully processed the latest form response", "Success");
  } catch (error) {
    Logger.log("Error processing form data: " + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast("Error processing form data: " + error.toString(), "Error");
  }
}