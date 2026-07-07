/**
 * Setlist.fm integration for Concert Tracker
 * This script adds setlist data to the Artist Breakdown analysis
 */

// Constants
const SETLIST_API_BASE = "https://api.setlist.fm/rest/1.0";
const SETLIST_COLUMN_NAME = "Setlist";

// Store API key in Script Properties for security
// To set up: In Apps Script, go to Project Settings > Script Properties and add:
// - Property Name: SETLIST_API_KEY
// - Value: Your actual API key
function getSetlistApiKey() {
  try {
    return PropertiesService.getScriptProperties().getProperty("SETLIST_API_KEY") || null;
  } catch(e) {
    Logger.log("Error accessing Script Properties for SETLIST_API_KEY.");
    return null;
  }
}

/**
 * Parse a setlist.fm URL to extract the setlist ID
 * @param {string} setlistUrl - The setlist.fm URL
 * @return {string|null} The extracted setlist ID or null if invalid
 */
function parseSetlistUrl(setlistUrl) {
  if (!setlistUrl || typeof setlistUrl !== 'string') {
    Logger.log("Invalid URL provided to parser");
    return null;
  }
  
  try {
    // Try multiple patterns to handle different URL formats
    
    // Format 1: https://www.setlist.fm/setlist/artist-name/date/venue-location-setlistId.html
    let match = setlistUrl.match(/\/setlist\/[^/]+\/[^/]+\/[^/]+-([a-f0-9]+)\.html/i);
    if (match && match[1]) {
      return match[1]; // Return the hexadecimal ID directly
    }
    
    // Format 2: https://www.setlist.fm/setlist/setlistId.html
    match = setlistUrl.match(/\/setlist\/([a-f0-9]+)\.html/i);
    if (match && match[1]) {
      return match[1]; // Return the hexadecimal ID directly
    }
    
    // New format: get the last segment before .html and extract hexadecimal ID
    match = setlistUrl.match(/\/([^/]+)\.html/);
    if (match && match[1]) {
      // Extract the hex ID (last segment after the final dash)
      const segments = match[1].split('-');
      const potentialId = segments[segments.length - 1];
      
      // Verify it looks like a hex ID (should be mostly hex characters)
      if (/^[a-f0-9]+$/i.test(potentialId)) {
        return potentialId;
      }
    }
    
    Logger.log(`Could not extract setlist ID from URL: ${setlistUrl}`);
    return null;
  } catch (error) {
    Logger.log(`Error parsing setlist URL: ${error.toString()}`);
    return null;
  }
}

/**
 * Get setlist data from a URL with improved error handling
 * @param {string} setlistUrl - The setlist.fm URL
 * @return {Object} The setlist data or null if error
 */
function getSetlistData(setlistUrl) {
  try {
    // Log the attempt
    Logger.log(`Attempting to fetch setlist data from URL: ${setlistUrl}`);
    
    // Extract the setlist ID from the URL
    const setlistId = parseSetlistUrl(setlistUrl);
    
    // Final check - if we couldn't extract the ID, log and return
    if (!setlistId) {
      Logger.log(`Could not extract valid setlist ID from URL: ${setlistUrl}`);
      return null;
    }
    
    // Log the extracted ID
    Logger.log(`Extracted setlist ID: ${setlistId}`);
    
    // Construct the API URL - use the ID directly without transformation
    const apiUrl = `${SETLIST_API_BASE}/setlist/${setlistId}`;
    Logger.log(`Calling API URL: ${apiUrl}`);
    
    // Make the API request with proper error handling
    const apiKey = getSetlistApiKey();
    if (!apiKey) {
      Logger.log("Setlist API key not configured. Set SETLIST_API_KEY in Script Properties.");
      return null;
    }

    const options = {
      headers: {
        "x-api-key": apiKey,
        "Accept": "application/json"
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(apiUrl, options);
    const statusCode = response.getResponseCode();
    
    // Enhanced status code handling
    if (statusCode !== 200) {
      // Log specific error based on status code
      switch(statusCode) {
        case 401:
        case 403:
          Logger.log("Authentication error: API key is invalid or expired");
          break;
        case 404:
          Logger.log(`Setlist not found with ID: ${setlistId}`);
          break;
        case 429:
          Logger.log("Rate limit exceeded. Try again later.");
          break;
        default:
          Logger.log(`API request failed with status: ${statusCode}`);
      }
      
      Logger.log(`Response: ${response.getContentText().substring(0, 200)}...`);
      return null;
    }
    
    // Parse the JSON response
    const setlistData = JSON.parse(response.getContentText());
    Logger.log(`Successfully fetched setlist data for ID: ${setlistId}`);
    return setlistData;
    
  } catch (error) {
    Logger.log(`Error fetching setlist data: ${error.toString()}`);
    return null;
  }
}

/**
 * Process setlist data to extract song information
 * @param {Object} setlistData - The raw setlist data
 * @return {Object} Processed song data and statistics
 */
function processSetlistData(setlistData) {
  try {
    // Log the data we're processing
    Logger.log(`Processing setlist data: ${JSON.stringify(setlistData).substring(0, 100)}...`);
    
    if (!setlistData || !setlistData.sets) {
      Logger.log("No valid setlist data structure found");
      return { 
        success: false, 
        message: "No setlist data available" 
      };
    }
    
    // Extract event information
    const eventDate = setlistData.eventDate || "Unknown Date";
    const venueName = setlistData.venue ? setlistData.venue.name : "Unknown Venue";
    const cityName = setlistData.venue && setlistData.venue.city ? setlistData.venue.city.name : "Unknown City";
    const tourName = setlistData.tour ? setlistData.tour.name : "";
    
    // Extract songs from all sets
    const allSongs = [];
    const sets = setlistData.sets.set;
    
    // Handle empty setlist
    if (!sets || sets.length === 0) {
      Logger.log("No sets found in the setlist data");
      return {
        success: true,
        eventDate: eventDate,
        venue: venueName,
        city: cityName,
        tour: tourName,
        songs: [],
        totalSongs: 0,
        message: "No songs listed in this setlist"
      };
    }
    
    // Process each set in the setlist
    sets.forEach((set, setIndex) => {
      if (set.song) {
        set.song.forEach((song, songIndex) => {
          // Add set information to each song
          allSongs.push({
            name: song.name,
            info: song.info || "",
            encore: set.encore ? true : false,
            encoreNumber: set.encore || 0,
            setIndex: setIndex,
            position: songIndex + 1,
            isCover: song.cover ? true : false,
            coverArtist: song.cover ? song.cover.name : ""
          });
        });
      }
    });
    
    Logger.log(`Processed ${allSongs.length} songs from the setlist`);
    
    return {
      success: true,
      eventDate: eventDate,
      venue: venueName,
      city: cityName,
      tour: tourName,
      songs: allSongs,
      totalSongs: allSongs.length
    };
    
  } catch (error) {
    Logger.log(`Error processing setlist data: ${error.toString()}`);
    return { 
      success: false, 
      message: `Error processing setlist data: ${error.toString()}` 
    };
  }
}

/**
 * Get all setlists for a specific artist with improved rate limiting
 * @param {Array} allShows - Array of all shows for the artist
 * @param {number} setlistCol - The column index for setlist URLs
 * @return {Array} Array of processed setlist data
 */
function getArtistSetlists(allShows, setlistCol) {
  // Check that the setlist column index is valid
  if (setlistCol === -1) {
    Logger.log("Setlist column not found in the tracker sheet");
    return [];
  }
  
  // Find shows with setlist URLs
  const showsWithSetlists = allShows.filter(show => show[setlistCol] && show[setlistCol].trim() !== "");
  
  // Log the starting process
  Logger.log(`Starting to fetch ${showsWithSetlists.length} setlists out of ${allShows.length} shows`);
  
  // Display progress to user
  SpreadsheetApp.getActiveSpreadsheet().toast(`Found ${showsWithSetlists.length} potential setlists to fetch`, "Processing");
  
  // Array to store all setlist data
  const artistSetlists = [];
  let successCount = 0;
  let failCount = 0;
  
  // Exit early if no setlists found
  if (showsWithSetlists.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("No setlist URLs found to process", "Complete");
    return [];
  }
  
  // Loop through shows with setlists with adaptive delay
  let currentDelay = 1000; // Start with 1 second delay
  
  for (let index = 0; index < showsWithSetlists.length; index++) {
    const show = showsWithSetlists[index];
    const setlistUrl = show[setlistCol];
    
    // Skip empty URLs
    if (!setlistUrl || setlistUrl.trim() === "") continue;
    
    // Update progress every few requests
    if (index % 2 === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Fetching setlist ${index+1} of ${showsWithSetlists.length}...`, 
        "Processing"
      );
    }
    
    Logger.log(`Processing setlist URL (${index+1}/${showsWithSetlists.length}): ${setlistUrl}`);
    
    // Apply delay to avoid rate limits
    Utilities.sleep(currentDelay);
    
    // Get and process the setlist data
    const rawSetlistData = getSetlistData(setlistUrl);
    
    if (rawSetlistData) {
      const processedSetlist = processSetlistData(rawSetlistData);
      
      if (processedSetlist.success) {
        // Add the original URL to the processed data
        processedSetlist.url = setlistUrl;
        artistSetlists.push(processedSetlist);
        successCount++;
        
        // Decrease delay for successful requests (but maintain minimum)
        currentDelay = Math.max(500, currentDelay - 100);
        
        Logger.log(`Successfully processed setlist #${successCount}`);
      } else {
        failCount++;
        Logger.log(`Failed to process setlist: ${processedSetlist.message}`);
      }
    } else {
      failCount++;
      Logger.log(`No data returned for setlist URL: ${setlistUrl}`);
      
      // Increase delay for failed requests (might be hitting rate limits)
      currentDelay += 500;
    }
  }
  
  // Final log
  Logger.log(`Completed setlist processing. ${successCount} setlists successfully processed, ${failCount} failed.`);
  
  // Notify user of completion
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Processed ${successCount} setlists successfully` + (failCount > 0 ? `, ${failCount} failed` : ""), 
    "Complete"
  );
  
  return artistSetlists;
}

/**
 * Analyze setlist data to extract statistics
 * @param {Array} allSetlists - Array of all processed setlists
 * @return {Object} Statistics and song frequencies
 */
function analyzeSetlistData(allSetlists) {
  // If no setlists, return empty stats
  if (!allSetlists || allSetlists.length === 0) {
    Logger.log("No setlists to analyze");
    return {
      totalSetlists: 0,
      totalSongs: 0,
      uniqueSongs: 0,
      avgSongsPerShow: 0,
      songFrequency: [],
      mostPlayedSongs: []
    };
  }
  
  Logger.log(`Analyzing ${allSetlists.length} setlists`);
  
  // Count total songs and unique songs
  let totalSongs = 0;
  const songCounts = {};
  const songDetails = {};
  
  // Process each setlist
  allSetlists.forEach(setlist => {
    if (!setlist.songs) return;
    
    totalSongs += setlist.songs.length;
    
    // Process each song
    setlist.songs.forEach(song => {
      const songName = song.name;
      
      // Count song frequency
      if (!songCounts[songName]) {
        songCounts[songName] = 0;
        songDetails[songName] = {
          name: songName,
          count: 0,
          asEncore: 0,
          isCover: song.isCover,
          coverArtist: song.coverArtist || ""
        };
      }
      
      songCounts[songName]++;
      songDetails[songName].count++;
      
      // Count when played as encore
      if (song.encore) {
        songDetails[songName].asEncore++;
      }
    });
  });
  
  // Convert to array and sort by frequency
  const songFrequency = Object.keys(songCounts).map(songName => {
    return {
      song: songName,
      count: songCounts[songName],
      percentage: Math.round((songCounts[songName] / allSetlists.length) * 100),
      asEncore: songDetails[songName].asEncore,
      isCover: songDetails[songName].isCover,
      coverArtist: songDetails[songName].coverArtist
    };
  });
  
  // Sort by count (descending)
  songFrequency.sort((a, b) => b.count - a.count);
  
  // Calculate statistics
  const stats = {
    totalSetlists: allSetlists.length,
    totalSongs: totalSongs,
    uniqueSongs: Object.keys(songCounts).length,
    avgSongsPerShow: Math.round((totalSongs / allSetlists.length) * 10) / 10,
    songFrequency: songFrequency,
    mostPlayedSongs: songFrequency.slice(0, 20) // Top 20 songs
  };
  
  Logger.log(`Analysis complete: Found ${stats.uniqueSongs} unique songs across ${stats.totalSetlists} setlists`);
  
  return stats;
}

/**
 * Add setlist analysis to the artist breakdown sheet
 * @param {Sheet} artistSheet - The artist breakdown sheet
 * @param {number} startRow - The row to start adding content
 * @param {Object} setlistStats - The setlist statistics
 * @return {number} The number of rows added
 */
function addSetlistAnalysisToSheet(artistSheet, startRow, setlistStats) {
  // Get the color scheme from ARTIST_COLORS if available, or fallback to defaults
  const headerBg = typeof ARTIST_COLORS !== 'undefined' ? ARTIST_COLORS.HEADER_BG : "#0d47a1";
  const headerText = typeof ARTIST_COLORS !== 'undefined' ? ARTIST_COLORS.HEADER_TEXT : "#FFFFFF";
  const altRowColor = typeof ARTIST_COLORS !== 'undefined' ? ARTIST_COLORS.ALT_ROW_COLOR : "#F0F8FF";
  
  // If no setlists are available, add a message and return
  if (!setlistStats || setlistStats.totalSetlists === 0) {
    artistSheet.getRange(startRow, 1).setValue("Setlist Analysis").setFontWeight("bold");
    artistSheet.getRange(startRow + 1, 1).setValue("No setlist data available");
    return 2;
  }
  
  // Add the Setlist Analysis header
  const headerRange = artistSheet.getRange(startRow, 1, 1, 5);
  headerRange.merge()
    .setValue("Setlist Analysis")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(headerBg)  // Blue header using imported or default colors
    .setFontColor(headerText);  // White text
  
  let currentRow = startRow + 1;
  
  // Add setlist summary statistics
  artistSheet.getRange(currentRow, 1, 4, 2).setValues([
    ["Setlists Analyzed:", setlistStats.totalSetlists],
    ["Unique Songs:", setlistStats.uniqueSongs],
    ["Total Songs Played:", setlistStats.totalSongs],
    ["Avg. Songs Per Show:", setlistStats.avgSongsPerShow]
  ]);
  
  artistSheet.getRange(currentRow, 1, 4, 1).setFontWeight("bold");
  currentRow += 5;  // Add an extra row of spacing
  
  // Add most played songs
  if (setlistStats.mostPlayedSongs && setlistStats.mostPlayedSongs.length > 0) {
    // Add most played songs header with blue theme formatting
    const songsHeaderRange = artistSheet.getRange(currentRow, 1, 1, 4);
    songsHeaderRange.merge()
      .setValue("Most Played Songs")
      .setFontWeight("bold")
      .setBackground(headerBg)
      .setFontColor(headerText)
      .setHorizontalAlignment("center");
    currentRow++;
    
    // Add column headers
    const headerValues = [["Song", "Times Played", "% of Shows", "Notes"]];
    artistSheet.getRange(currentRow, 1, 1, 4).setValues(headerValues)
      .setFontWeight("bold")
      .setBackground("#e0e0e0");
    currentRow++;
    
    // Add song data
    setlistStats.mostPlayedSongs.forEach((song, index) => {
      let notes = "";
      
      // Add cover information if applicable
      if (song.isCover) {
        notes += `Cover (${song.coverArtist}) `;
      }
      
      // Add encore information if applicable
      if (song.asEncore > 0) {
        const encorePercent = Math.round((song.asEncore / song.count) * 100);
        notes += `Encore ${song.asEncore}x (${encorePercent}%)`;
      }
      
      // Add the song data
      artistSheet.getRange(currentRow, 1, 1, 4).setValues([
        [song.song, song.count, `${song.percentage}%`, notes]
      ]);
      
      // Add alternate row coloring for readability
      if (index % 2 === 1) {
        artistSheet.getRange(currentRow, 1, 1, 4).setBackground(altRowColor);
      }
      
      currentRow++;
    });
    
    // Add a note about the data source
    currentRow++;
    artistSheet.getRange(currentRow, 1, 1, 4).merge()
      .setValue("Data source: setlist.fm API")
      .setFontStyle("italic")
      .setFontSize(9)
      .setHorizontalAlignment("center");
    currentRow++;
  }
  
  // Return the number of rows added
  return currentRow - startRow;
}

/**
 * Search setlist.fm for a setlist matching artist + date.
 * Tries with venue name first for precision, falls back to artist + date only.
 * Returns the setlist URL string, or null if not found.
 */
function searchSetlistFm(artistName, dateValue, venueName) {
  const apiKey = getSetlistApiKey();
  if (!apiKey) {
    Logger.log("Setlist API key not configured — skipping auto-search");
    return null;
  }

  // Format date as dd-MM-yyyy (setlist.fm's required format)
  let formattedDate = "";
  try {
    let d = dateValue;
    if (!(d instanceof Date)) {
      if (typeof d === "string" && d.includes('/')) {
        const p = d.split('/');
        if (p.length === 3) d = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
      } else {
        d = new Date(d);
      }
    }
    if (!d || isNaN(d.getTime())) {
      Logger.log("Invalid date for setlist search");
      return null;
    }
    formattedDate = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy");
  } catch (e) {
    Logger.log("Error formatting date for setlist search: " + e);
    return null;
  }

  const baseParams = [
    `artistName=${encodeURIComponent(artistName)}`,
    `date=${formattedDate}`
  ];

  const options = {
    headers: { "x-api-key": apiKey, "Accept": "application/json" },
    muteHttpExceptions: true
  };

  function doSearch(params) {
    const url = `${SETLIST_API_BASE}/search/setlists?${params.join('&')}`;
    Logger.log(`Searching setlist.fm: ${url}`);
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() !== 200) return null;
      const data = JSON.parse(response.getContentText());
      if (data.setlist && data.setlist.length > 0) return data.setlist[0].url || null;
    } catch (e) {
      Logger.log("Error during setlist search: " + e);
    }
    return null;
  }

  // Try with venue first (more precise), then fall back to artist + date
  if (venueName) {
    const result = doSearch([...baseParams, `venueName=${encodeURIComponent(venueName)}`]);
    if (result) return result;
    Utilities.sleep(500);
  }

  return doSearch(baseParams);
}

/**
 * Auto-fetch and store a setlist URL for the given tracker row if one isn't already set.
 * Returns the URL if found, null otherwise.
 */
function autoFetchSetlistUrl(sheet, row, headers) {
  const setlistCol = findSetlistColumn(headers);
  if (setlistCol === -1) return null;

  // Don't overwrite an existing URL
  const existing = sheet.getRange(row, setlistCol + 1).getValue();
  if (existing && existing.toString().trim() !== "") return existing;

  const artistCol = headers.indexOf("Artist");
  const dateCol   = headers.indexOf("Date");
  const venueCol  = headers.indexOf("Venue");

  const artistName = artistCol !== -1 ? sheet.getRange(row, artistCol + 1).getValue() : null;
  const dateValue  = dateCol   !== -1 ? sheet.getRange(row, dateCol   + 1).getValue() : null;
  const venueName  = venueCol  !== -1 ? sheet.getRange(row, venueCol  + 1).getValue() : null;

  if (!artistName || !dateValue) return null;

  const url = searchSetlistFm(artistName, dateValue, venueName);
  if (url) {
    sheet.getRange(row, setlistCol + 1).setValue(url);
    Logger.log(`Auto-stored setlist URL for ${artistName}: ${url}`);
  }
  return url;
}

/**
 * Find the setlist column index with improved reliability
 * @param {Array} headers - Headers from main tracker sheet
 * @return {number} The column index for setlist URLs (-1 if not found)
 */
function findSetlistColumn(headers) {
  // Try multiple approaches to find the setlist column
  
  // 1. Try exact match for "Setlist"
  let setlistCol = headers.indexOf(SETLIST_COLUMN_NAME);
  if (setlistCol !== -1) {
    Logger.log(`Found setlist column "${SETLIST_COLUMN_NAME}" at index ${setlistCol}`);
    return setlistCol;
  }
  
  // 2. Try case-insensitive match
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && headers[i].toString().toLowerCase() === SETLIST_COLUMN_NAME.toLowerCase()) {
      Logger.log(`Found setlist column with case-insensitive match at index ${i}`);
      return i;
    }
  }
  
  // 3. Try to find any column containing "setlist" in the name
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && headers[i].toString().toLowerCase().includes("setlist")) {
      Logger.log(`Found column with partial match "${headers[i]}" at index ${i}`);
      return i;
    }
  }
  
  // 4. Fallback to column N (index 13) if it exists
  if (headers.length > 13) {
    Logger.log(`Using fallback column N (index 13) for setlists`);
    return 13;
  }
  
  // No column found
  Logger.log("Could not find setlist column");
  return -1;
}

/**
 * Debug function to log the setlist columns from all artist shows
 * This helps identify if the correct column is being used
 */
function debugSetlistColumn(artistShows, headers) {
  const setlistCol = findSetlistColumn(headers);
  
  Logger.log(`Setlist column name: "${SETLIST_COLUMN_NAME}"`);
  Logger.log(`Setlist column index: ${setlistCol}`);
  Logger.log(`Headers: ${JSON.stringify(headers)}`);
  
  // Log the values in the supposed setlist column
  if (setlistCol !== -1) {
    Logger.log("First 5 values in setlist column:");
    for (let i = 0; i < Math.min(5, artistShows.length); i++) {
      Logger.log(`  [${i+1}] ${artistShows[i][setlistCol]}`);
    }
  }
}

/**
 * Main function to integrate setlist data in artist breakdown
 * @param {Sheet} artistSheet - The artist breakdown sheet
 * @param {Array} artistShows - Shows for the selected artist
 * @param {number} row - Current row in artist sheet
 * @param {Array} headers - Headers from main tracker sheet
 * @return {number} Updated row position
 */
function integrateSetlistData(artistSheet, artistShows, row, headers) {
  // Get the color scheme from ARTIST_COLORS if available, or fallback to defaults
  const headerBg = typeof ARTIST_COLORS !== 'undefined' ? ARTIST_COLORS.HEADER_BG : "#0d47a1";
  const headerText = typeof ARTIST_COLORS !== 'undefined' ? ARTIST_COLORS.HEADER_TEXT : "#FFFFFF";
  
  // Debug the setlist column
  debugSetlistColumn(artistShows, headers);
  
  // Find the setlist column index with improved reliability
  const setlistCol = findSetlistColumn(headers);
  
  // Check if any show has a setlist URL
  const hasSetlists = artistShows.some(show => 
    show[setlistCol] && typeof show[setlistCol] === 'string' && show[setlistCol].trim() !== "" && 
    show[setlistCol].toLowerCase().includes("setlist.fm")
  );
  
  // If no setlists, return the current row without changes
  if (setlistCol === -1 || !hasSetlists) {
    Logger.log(`No valid setlists found for this artist (setlistCol: ${setlistCol}, hasSetlists: ${hasSetlists})`);
    
    // Add a notice but with proper formatting
    row += 2;
    const headerRange = artistSheet.getRange(row, 1, 1, 3);
    headerRange.merge()
      .setValue("Setlist Analysis")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground(headerBg)
      .setFontColor(headerText)
      .setHorizontalAlignment("center");
    row++;
    
    if (setlistCol === -1) {
      artistSheet.getRange(row, 1).setValue("Setlist column not found. Add a column named 'Setlist' to your tracker to use this feature.");
    } else {
      artistSheet.getRange(row, 1).setValue("No setlist.fm URLs found for this artist.");
    }
    row += 2;
    
    return row;
  }
  
  // Add some spacing before the setlist section
  row += 2;
  
  // Get all setlists for the artist
  const allSetlists = getArtistSetlists(artistShows, setlistCol);
  
  // Analyze the setlist data
  const setlistStats = analyzeSetlistData(allSetlists);
  
  // Add the analysis to the sheet
  const rowsAdded = addSetlistAnalysisToSheet(artistSheet, row, setlistStats);
  
  // Return the updated row position
  return row + rowsAdded;
}