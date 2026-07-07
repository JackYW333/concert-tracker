function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Concert Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Returns the current cache generation number, used to bust stale entries.
function getCacheGeneration() {
  return PropertiesService.getScriptProperties().getProperty('CACHE_GEN') || '0';
}

// Increment the generation so all existing cache entries are ignored on the next request.
function incrementCacheGeneration() {
  const props = PropertiesService.getScriptProperties();
  const next = parseInt(props.getProperty('CACHE_GEN') || '0') + 1;
  props.setProperty('CACHE_GEN', next.toString());
}

// Wraps a function call with CacheService. Returns cached result if available,
// otherwise calls fn(), stores the result, and returns it.
// ttl is in seconds (max 21600 = 6 hours). Silently skips caching if the
// serialised result exceeds CacheService's 100 KB per-entry limit.
function withCache(key, ttl, fn) {
  const fullKey = 'ct_' + getCacheGeneration() + '_' + key;
  const cache = CacheService.getScriptCache();
  const hit = cache.get(fullKey);
  if (hit) {
    try { return JSON.parse(hit); } catch (_) {}
  }
  const result = fn();
  try { cache.put(fullKey, JSON.stringify(result), ttl); } catch (_) {}
  return result;
}

function handleApiRequest(e) {
  const action = e.parameter.action;
  const TTL = 3600; // 1 hour
  let result;
  try {
    switch (action) {
      case 'getArtists':
        result = withCache('artists', TTL, getArtists); break;
      case 'getArtistData': {
        const artist = e.parameter.artist || '';
        result = withCache('artist_' + artist, TTL, () => getArtistData(artist)); break;
      }
      case 'getAnalysis':
        result = withCache('analysis', TTL, getAnalysisData); break;
      case 'getVenues':
        result = withCache('venues', TTL, getVenues); break;
      case 'getVenueData': {
        const venue = e.parameter.venue || '';
        result = withCache('venue_' + venue, TTL, () => getVenueData(venue)); break;
      }
      case 'getEventData':
        result = getEventData(e.parameter.date, e.parameter.venue); break;
      case 'getYears':
        result = withCache('years', TTL, getYears); break;
      case 'getYearData': {
        const year = e.parameter.year || '';
        result = withCache('year_' + year, TTL, () => getYearData(year)); break;
      }
      case 'getUpcomingShows':
        // Short TTL because daysUntil is calculated server-side
        result = withCache('upcoming', 600, getUpcomingShows); break;
      case 'getCompanionData': {
        const companion = e.parameter.companion || '';
        result = withCache('companion_' + companion, TTL, () => getCompanionData(companion)); break;
      }
      case 'getCityData': {
        const city = e.parameter.city || '';
        result = withCache('city_' + city, TTL, () => getCityData(city)); break;
      }
      default: result = { error: 'Unknown action', success: false };
    }
  } catch (err) {
    result = { error: err.toString(), success: false };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Get list of all artists from the Tracker sheet
function getArtists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  
  if (!mainSheet) {
    return { error: "Tracker sheet not found", success: false };
  }
  
  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  const artistCol = headers.indexOf("Artist");
  const statusCol = headers.indexOf("Status");
  
  if (artistCol === -1) {
    return { error: "Artist column not found", success: false };
  }
  
  // Get all unique artists that have at least one attended show
  const attendedShows = data.slice(1).filter(row => row[statusCol] === "Attended");
  const uniqueArtists = [...new Set(attendedShows.map(row => row[artistCol]))].sort();
  
  return { 
    artists: uniqueArtists,
    success: true
  };
}

// Get all data for a specific artist
function getArtistData(artistName) {
  if (!artistName) {
    return { error: "No artist specified", success: false };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  
  if (!mainSheet) {
    return { error: "Tracker sheet not found", success: false };
  }
  
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
  
  // Filter for artist's shows
  const artistShows = data.slice(1).filter(row => 
    row[artistCol] === artistName && row[statusCol] === "Attended"
  );
  
  // Create a map of events to their headliners
  const eventHeadliners = {};
  
  // First pass to identify headliners for each event
  data.slice(1).forEach(show => {
    // Skip non-attended shows
    if (show[statusCol] !== "Attended") return;
    
    const role = show[roleCol] || "";
    const isHeadliner = role.toLowerCase().includes("headline") || role.toLowerCase().includes("headliner");
    
    if (isHeadliner) {
      const showArtist = show[artistCol];
      const showDate = show[dateCol];
      const venue = show[venueCol] || "";
      
      // Create a unique event key using date and venue
      let eventKey = "";
      if (showDate instanceof Date) {
        eventKey = `${Utilities.formatDate(showDate, Session.getScriptTimeZone(), "yyyy-MM-dd")}_${venue}`;
      } else if (typeof showDate === 'string') {
        eventKey = `${showDate}_${venue}`;
      }
      
      if (!eventHeadliners[eventKey]) {
        eventHeadliners[eventKey] = [];
      }
      // Add this artist as a headliner
      eventHeadliners[eventKey].push(showArtist);
    }
  });
  
  // Format shows for response
  const concerts = artistShows.map(show => {
    // Format date
    let dateStr = "";
    if (show[dateCol] instanceof Date) {
      dateStr = Utilities.formatDate(show[dateCol], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof show[dateCol] === "string") {
      // Try to parse the date string
      if (show[dateCol].includes('/')) {
        const dateParts = show[dateCol].split('/');
        if (dateParts.length === 3) {
          // Try to handle dd/mm/yyyy format
          const day = parseInt(dateParts[0], 10);
          const month = parseInt(dateParts[1], 10) - 1;
          const year = parseInt(dateParts[2], 10);
          dateStr = `${year}-${(month+1).toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}`;
        }
      }
    }
    
    // Process companions
    let companions = [];
    if (withCol >= 0 && show[withCol]) {
      companions = show[withCol].split(',').map(c => c.trim()).filter(c => c !== "");
    }
    
    // Determine headliner for support acts
    let headliner = "";
    if (roleCol >= 0 && show[roleCol] && show[roleCol].toLowerCase().includes("support")) {
      // Get event key for headliner lookup
      let eventKey = "";
      if (show[dateCol] instanceof Date) {
        eventKey = `${Utilities.formatDate(show[dateCol], Session.getScriptTimeZone(), "yyyy-MM-dd")}_${show[venueCol] || ""}`;
      } else if (typeof show[dateCol] === 'string') {
        eventKey = `${show[dateCol]}_${show[venueCol] || ""}`;
      }
      
      const eventHeadlinersList = eventHeadliners[eventKey] || [];
      // Filter out the current artist (in case they're incorrectly listed)
      const otherHeadliners = eventHeadlinersList.filter(artist => artist !== artistName);
      
      if (otherHeadliners.length > 0) {
        headliner = otherHeadliners.join(" & ");
      }
    }
    
    return {
      date: dateStr,
      venue: show[venueCol] || "",
      city: show[cityCol] || "",
      role: show[roleCol] || "",
      eventName: show[eventNameCol] || "",
      headliner: headliner,
      rating: show[ratingCol] || "",
      notes: show[notesCol] || "",
      companions: companions
    };
  });
  
  // Sort concerts by date
  concerts.sort((a, b) => a.date.localeCompare(b.date));
  
  return { 
    artist: artistName,
    concerts: concerts,
    success: true
  };
}

// Get data for the analysis dashboard
function getAnalysisData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  const analysisSheet = ss.getSheetByName("Concert Analysis");
  
  if (!mainSheet) {
    return { error: "Tracker sheet not found", success: false };
  }
  
  // Get all concert data
  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indexes
  const artistCol = headers.indexOf("Artist");
  const dateCol = headers.indexOf("Date");
  const venueCol = headers.indexOf("Venue");
  const cityCol = headers.indexOf("City");
  const withCol = headers.indexOf("Who I Went With");
  const statusCol = headers.indexOf("Status");
  const ratingCol = headers.indexOf("Rating");
  
  // Filter for attended shows
  const attendedShows = data.slice(1).filter(row => row[statusCol] === "Attended");
  
  // Group by events (same date + venue = same event)
  const events = {};
  
  // Artist counter
  const artistCounts = {};
  const artistsByYear = {};
  
  // Count total artist appearances (not unique)
  let totalArtistAppearances = 0;
  
  // Process all attended shows
  attendedShows.forEach(show => {
    // Artist count
    const artist = show[artistCol];
    if (artist) {
      artistCounts[artist] = (artistCounts[artist] || 0) + 1;
      totalArtistAppearances++; // Count every artist appearance
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
  
  // Count venues
  const venueCount = {};
  Object.values(events).forEach(event => {
    if (event.venue) {
      venueCount[event.venue] = (venueCount[event.venue] || 0) + 1;
    }
  });
  const topVenues = Object.entries(venueCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([name, count]) => ({ name, count }));
  
  // Count cities
  const cityCount = {};
  Object.values(events).forEach(event => {
    if (event.city) {
      cityCount[event.city] = (cityCount[event.city] || 0) + 1;
    }
  });
  const topCities = Object.entries(cityCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, count]) => ({ name, count }));
  
  // Count artists
  const topArtists = Object.entries(artistCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([name, count]) => ({ name, count }));
  
  // Count events by year
  const eventsByYear = {};
  Object.values(events).forEach(event => {
    if (event.year) {
      eventsByYear[event.year] = (eventsByYear[event.year] || 0) + 1;
    }
  });
  
  // Sort years
  const years = Object.keys(eventsByYear).sort();
  
  // Count companions
  const companionCount = {};
  Object.values(events).forEach(event => {
    event.companions.forEach(companion => {
      companionCount[companion] = (companionCount[companion] || 0) + 1;
    });
  });
  const topCompanions = Object.entries(companionCount)
    .sort((a, b) => b[1] - a[1])
    .filter(([_, count]) => count >= 2) // Only show companions with at least 2 events
    .slice(0, 10)
    .map(([name, count]) => ({ name, count }));
  
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
  
  // Format data for year charts
  const eventsYearData = years.map(year => ({
    year: parseInt(year),
    count: eventsByYear[year]
  }));
  
  // Artist count by year data
  const artistYearData = years.map(year => ({
    year: parseInt(year),
    count: artistsByYear[year] ? artistsByYear[year].size : 0
  }));
  
  return {
    summary: {
      totalEvents: Object.keys(events).length,
      totalArtists: totalArtistAppearances, // NEW: Total artist appearances
      uniqueArtists: Object.keys(artistCounts).length,
      uniqueVenues: Object.keys(venueCount).length,
      uniqueCities: Object.keys(cityCount).length,
      avgRating: avgRating
    },
    topArtists: topArtists,
    topVenues: topVenues,
    topCities: topCities,
    eventsYearData: eventsYearData,
    artistYearData: artistYearData,
    topCompanions: topCompanions,
    success: true
  };
}

// Client-ready functions that will be called from the web interface
function getArtistsForClient() {
  return getArtists();
}

function getArtistDataForClient(artist) {
  return getArtistData(artist);
}

function getAnalysisForClient() {
  return getAnalysisData();
}

/**
 * Get all venues from the Tracker sheet
 * @return {Object} - Object with venues array or error
 */
function getVenues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  
  if (!mainSheet) {
    return { error: "Tracker sheet not found", success: false };
  }
  
  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  const venueCol = headers.indexOf("Venue");
  const statusCol = headers.indexOf("Status");
  
  if (venueCol === -1) {
    return { error: "Venue column not found", success: false };
  }
  
  // Get all unique venues that have at least one attended show
  const attendedShows = data.slice(1).filter(row => row[statusCol] === "Attended");
  const uniqueVenues = [...new Set(attendedShows.map(row => row[venueCol]))].sort();
  
  return { 
    venues: uniqueVenues,
    success: true
  };
}
/**
 * Get all data for a specific venue - Updated to only show headliner ratings
 * @param {string} venueName - Name of the venue to analyze
 * @return {Object} - Object containing venue data or error
 */
function getVenueData(venueName) {
  if (!venueName) {
    return { error: "No venue specified", success: false };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  
  if (!mainSheet) {
    return { error: "Tracker sheet not found", success: false };
  }
  
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
  
  // Filter for venue's shows
  const venueShows = data.slice(1).filter(row => 
    row[venueCol] === venueName && row[statusCol] === "Attended"
  );
  
  if (venueShows.length === 0) {
    return {
      venue: venueName,
      events: [],
      summary: {
        totalEvents: 0,
        totalArtists: 0,
        uniqueArtists: 0,
        city: "Unknown",
        avgRating: "N/A",
        firstVisit: null,
        lastVisit: null,
        timeSpan: "N/A"
      },
      success: true
    };
  }
  
  // Group by dates to count unique events and consolidate artists
  const eventsMap = {};
  
  venueShows.forEach(show => {
    // Create date key
    let dateKey;
    if (show[dateCol] instanceof Date) {
      dateKey = Utilities.formatDate(show[dateCol], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      dateKey = show[dateCol].toString();
    }
    
    const artist = show[artistCol];
    const eventName = show[eventNameCol] || "";
    const companions = show[withCol] ? show[withCol].split(',').map(c => c.trim()) : [];
    const role = show[roleCol] || "";
    const rating = show[ratingCol] || "";
    
    if (!eventsMap[dateKey]) {
      eventsMap[dateKey] = {
        date: show[dateCol],
        eventName: eventName,
        artists: [],
        roles: {},
        ratings: {}, // Store all ratings by artist
        companions: companions,
        notes: show[notesCol] || ""
      };
    }
    
    // Add artist if not already added
    if (artist && !eventsMap[dateKey].artists.includes(artist)) {
      eventsMap[dateKey].artists.push(artist);
      eventsMap[dateKey].roles[artist] = role;
      eventsMap[dateKey].ratings[artist] = rating;
    }
    
    // Merge companions without duplicates
    companions.forEach(companion => {
      if (companion && !eventsMap[dateKey].companions.includes(companion)) {
        eventsMap[dateKey].companions.push(companion);
      }
    });
  });
  
  // Convert to array and sort by date
  const events = Object.values(eventsMap).map(event => {
    // Format date
    let dateStr = "";
    if (event.date instanceof Date) {
      dateStr = Utilities.formatDate(event.date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof event.date === "string") {
      // Try to parse the date string
      if (event.date.includes('/')) {
        const dateParts = event.date.split('/');
        if (dateParts.length === 3) {
          // Try to handle dd/mm/yyyy format
          const day = parseInt(dateParts[0], 10);
          const month = parseInt(dateParts[1], 10) - 1;
          const year = parseInt(dateParts[2], 10);
          dateStr = `${year}-${(month+1).toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}`;
        }
      } else {
        dateStr = event.date;
      }
    }
    
    // Format artists with roles
    const artistsWithRoles = event.artists.map(artist => {
      const role = event.roles[artist] || "";
      return {
        name: artist,
        role: role
      };
    });
    
    // Find headliner rating - try multiple approaches
    let headlinerRating = "";
    
    // Method 1: Look for explicit headliner roles
    for (const artist of event.artists) {
      const role = event.roles[artist] || "";
      const rating = event.ratings[artist] || "";
      
      // Check for various headliner indicators
      if (role && (
        role.toLowerCase().includes("headline") || 
        role.toLowerCase().includes("headliner") ||
        role.toLowerCase() === "headline" ||
        role.toLowerCase() === "headliner"
      )) {
        if (rating && rating !== "" && !isNaN(parseFloat(rating))) {
          headlinerRating = rating;
          break;
        }
      }
    }
    
    // Method 2: If no explicit headliner found, look for non-support acts with ratings
    if (!headlinerRating) {
      for (const artist of event.artists) {
        const role = event.roles[artist] || "";
        const rating = event.ratings[artist] || "";
        
        // If role doesn't indicate support, consider it a potential headliner
        if (!role.toLowerCase().includes("support") && 
            !role.toLowerCase().includes("opening") &&
            rating && rating !== "" && !isNaN(parseFloat(rating))) {
          headlinerRating = rating;
          break;
        }
      }
    }
    
    // Method 3: If still no rating found, use the first available rating
    if (!headlinerRating) {
      for (const artist of event.artists) {
        const rating = event.ratings[artist] || "";
        if (rating && rating !== "" && !isNaN(parseFloat(rating))) {
          headlinerRating = rating;
          break;
        }
      }
    }
    
    return {
      date: dateStr,
      eventName: event.eventName,
      artists: artistsWithRoles,
      companions: event.companions,
      rating: headlinerRating,
      notes: event.notes
    };
  });
  
  // Sort events by date
  events.sort((a, b) => a.date.localeCompare(b.date));
  
  // Count artists
  const allArtists = [];
  const artistsSet = new Set();
  
  venueShows.forEach(show => {
    const artist = show[artistCol];
    if (artist) {
      allArtists.push(artist);
      artistsSet.add(artist);
    }
  });
  
  // Get the city (should be the same for all shows at this venue)
  const city = venueShows[0][cityCol] || "Unknown";
  
  // Calculate average rating - only from events that have ratings
  const validRatings = events.map(event => event.rating)
    .filter(rating => rating !== "" && rating !== null && !isNaN(parseFloat(rating)));
  
  let avgRating = "N/A";
  if (validRatings.length > 0) {
    const sum = validRatings.reduce((acc, val) => acc + parseFloat(val), 0);
    avgRating = (sum / validRatings.length).toFixed(1);
  }
  
  // Calculate first and last visit dates
  const sortedShows = [...venueShows].sort((a, b) => {
    const dateA = a[dateCol] instanceof Date ? a[dateCol] : new Date(a[dateCol]);
    const dateB = b[dateCol] instanceof Date ? b[dateCol] : new Date(b[dateCol]);
    return dateA - dateB;
  });
  
  const firstVisit = sortedShows[0][dateCol];
  const lastVisit = sortedShows[sortedShows.length - 1][dateCol];
  
  // Format dates
  let firstVisitStr = firstVisit instanceof Date ? 
    Utilities.formatDate(firstVisit, Session.getScriptTimeZone(), "yyyy-MM-dd") : firstVisit;
  
  let lastVisitStr = lastVisit instanceof Date ? 
    Utilities.formatDate(lastVisit, Session.getScriptTimeZone(), "yyyy-MM-dd") : lastVisit;
  
  // Calculate time span
  let timeSpan = "N/A";
  try {
    if (firstVisit instanceof Date && lastVisit instanceof Date) {
      const diffTime = Math.abs(lastVisit - firstVisit);
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
  
  // Count top artists at this venue
  const artistFrequency = {};
  allArtists.forEach(artist => {
    artistFrequency[artist] = (artistFrequency[artist] || 0) + 1;
  });
  
  const topArtists = Object.entries(artistFrequency)
    .sort((a, b) => b[1] - a[1])
    .map(([name, count]) => ({ name, count }));
  
  // Count companions
  const companionFrequency = {};
  events.forEach(event => {
    event.companions.forEach(companion => {
      companionFrequency[companion] = (companionFrequency[companion] || 0) + 1;
    });
  });
  
  const topCompanions = Object.entries(companionFrequency)
    .sort((a, b) => b[1] - a[1])
    .map(([name, count]) => ({ name, count }));
  
  return { 
    venue: venueName,
    events: events,
    topArtists: topArtists,
    topCompanions: topCompanions,
    summary: {
      totalEvents: events.length,
      totalArtists: allArtists.length,
      uniqueArtists: artistsSet.size,
      city: city,
      avgRating: avgRating,
      firstVisit: firstVisitStr,
      lastVisit: lastVisitStr,
      timeSpan: timeSpan
    },
    success: true
  };
}
/**
 * Client-accessible wrapper for getVenues
 * @return {Object} - Venues data or error object
 */
function getVenuesForClient() {
  return getVenues();
}

/**
 * Client-accessible wrapper for getVenueData
 * @param {string} venue - Name of the venue to analyze
 * @return {Object} - Venue data or error object
 */
function getVenueDataForClient(venue) {
  return getVenueData(venue);
}

/**
 * Get all data for a specific city - venues, events, and top artists within it.
 * @param {string} cityName - Name of the city to analyze
 * @return {Object} - Object containing city data or error
 */
function getCityData(cityName) {
  if (!cityName) {
    return { error: "No city specified", success: false };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");

  if (!mainSheet) {
    return { error: "Tracker sheet not found", success: false };
  }

  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];

  const artistCol = headers.indexOf("Artist");
  const dateCol = headers.indexOf("Date");
  const venueCol = headers.indexOf("Venue");
  const cityCol = headers.indexOf("City");
  const roleCol = headers.indexOf("Role");
  const eventNameCol = headers.indexOf("Event Name");
  const ratingCol = headers.indexOf("Rating");
  const withCol = headers.indexOf("Who I Went With");
  const statusCol = headers.indexOf("Status");

  const cityShows = data.slice(1).filter(row =>
    row[cityCol] === cityName && row[statusCol] === "Attended"
  );

  if (cityShows.length === 0) {
    return {
      city: cityName,
      events: [],
      topArtists: [],
      topVenues: [],
      summary: {
        totalEvents: 0,
        totalArtists: 0,
        uniqueArtists: 0,
        uniqueVenues: 0,
        avgRating: "N/A",
        firstVisit: null,
        lastVisit: null,
        timeSpan: "N/A"
      },
      success: true
    };
  }

  // Group by date + venue to get unique events
  const eventsMap = {};
  cityShows.forEach(show => {
    let dateKey;
    if (show[dateCol] instanceof Date) {
      dateKey = Utilities.formatDate(show[dateCol], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      dateKey = show[dateCol].toString();
    }
    const venue = show[venueCol] || "";
    const key = `${dateKey}|${venue}`;
    const artist = show[artistCol];
    const eventName = show[eventNameCol] || "";
    const companions = show[withCol] ? show[withCol].split(',').map(c => c.trim()) : [];
    const role = show[roleCol] || "";
    const rating = show[ratingCol] || "";

    if (!eventsMap[key]) {
      eventsMap[key] = {
        date: show[dateCol],
        venue: venue,
        eventName: eventName,
        artists: [], roles: {}, ratings: {},
        companions: []
      };
    }
    if (artist && !eventsMap[key].artists.includes(artist)) {
      eventsMap[key].artists.push(artist);
      eventsMap[key].roles[artist] = role;
      eventsMap[key].ratings[artist] = rating;
    }
    companions.forEach(c => {
      if (c && !eventsMap[key].companions.includes(c)) eventsMap[key].companions.push(c);
    });
  });

  const events = Object.values(eventsMap).map(event => {
    let dateStr = "";
    if (event.date instanceof Date) {
      dateStr = Utilities.formatDate(event.date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof event.date === "string") {
      if (event.date.includes('/')) {
        const p = event.date.split('/');
        if (p.length === 3) dateStr = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
      } else {
        dateStr = event.date;
      }
    }

    const artistsWithRoles = event.artists.map(a => ({ name: a, role: event.roles[a] || "" }));

    let rating = "";
    for (const a of event.artists) {
      if ((event.roles[a] || "").toLowerCase().includes("headline") && event.ratings[a]) {
        rating = event.ratings[a]; break;
      }
    }
    if (!rating) {
      for (const a of event.artists) { if (event.ratings[a]) { rating = event.ratings[a]; break; } }
    }

    return {
      date: dateStr,
      venue: event.venue,
      eventName: event.eventName,
      artists: artistsWithRoles,
      companions: event.companions,
      rating: rating
    };
  });

  events.sort((a, b) => a.date.localeCompare(b.date));

  // Counts
  const artistCounts = {};
  cityShows.forEach(row => {
    const a = row[artistCol];
    if (a) artistCounts[a] = (artistCounts[a] || 0) + 1;
  });
  const venueCounts = {};
  events.forEach(ev => { if (ev.venue) venueCounts[ev.venue] = (venueCounts[ev.venue] || 0) + 1; });

  const topArtists = Object.entries(artistCounts).sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count }));
  const topVenues = Object.entries(venueCounts).sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count }));

  const ratings = events.map(e => parseFloat(e.rating)).filter(r => !isNaN(r));
  const avgRating = ratings.length > 0
    ? (ratings.reduce((a, b) => a + b, 0) / ratings.length).toFixed(1)
    : "N/A";

  // First/last visit + time span
  const sortedShows = [...cityShows].sort((a, b) => {
    const dA = a[dateCol] instanceof Date ? a[dateCol] : new Date(a[dateCol]);
    const dB = b[dateCol] instanceof Date ? b[dateCol] : new Date(b[dateCol]);
    return dA - dB;
  });
  const firstVisit = sortedShows[0][dateCol];
  const lastVisit = sortedShows[sortedShows.length - 1][dateCol];
  const firstVisitStr = firstVisit instanceof Date
    ? Utilities.formatDate(firstVisit, Session.getScriptTimeZone(), "yyyy-MM-dd") : firstVisit;
  const lastVisitStr = lastVisit instanceof Date
    ? Utilities.formatDate(lastVisit, Session.getScriptTimeZone(), "yyyy-MM-dd") : lastVisit;

  let timeSpan = "N/A";
  try {
    if (firstVisit instanceof Date && lastVisit instanceof Date) {
      const diffDays = Math.ceil(Math.abs(lastVisit - firstVisit) / 86400000);
      const diffYears = (diffDays / 365).toFixed(1);
      if (diffDays === 0) timeSpan = "Same day";
      else if (diffDays < 365) timeSpan = `${diffDays} day${diffDays === 1 ? "" : "s"}`;
      else timeSpan = `${diffYears} year${diffYears === "1.0" ? "" : "s"}`;
    }
  } catch (e) {}

  return {
    city: cityName,
    events: events,
    topArtists: topArtists,
    topVenues: topVenues,
    summary: {
      totalEvents: events.length,
      totalArtists: cityShows.length,
      uniqueArtists: Object.keys(artistCounts).length,
      uniqueVenues: Object.keys(venueCounts).length,
      avgRating: avgRating,
      firstVisit: firstVisitStr,
      lastVisit: lastVisitStr,
      timeSpan: timeSpan
    },
    success: true
  };
}

function getCityDataForClient(cityName) {
  return getCityData(cityName);
}

/**
 * Get all data for a specific event (identified by date + venue)
 * @param {string} dateStr - Date in yyyy-MM-dd format
 * @param {string} venueName - Name of the venue
 * @return {Object} - Event data or error
 */
function getEventData(dateStr, venueName) {
  if (!dateStr || !venueName) {
    return { error: "Date and venue are required", success: false };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  if (!mainSheet) return { error: "Tracker sheet not found", success: false };

  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];

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

  // Find all rows matching this date + venue
  const eventRows = data.slice(1).filter(row => {
    if (row[statusCol] !== "Attended") return false;

    let rowDateStr = "";
    if (row[dateCol] instanceof Date) {
      rowDateStr = Utilities.formatDate(row[dateCol], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof row[dateCol] === "string") {
      if (row[dateCol].includes('/')) {
        const p = row[dateCol].split('/');
        if (p.length === 3) {
          rowDateStr = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
        }
      } else {
        rowDateStr = row[dateCol];
      }
    }

    return rowDateStr === dateStr && row[venueCol] === venueName;
  });

  if (eventRows.length === 0) {
    return { error: "Event not found", success: false };
  }

  // Collect companions (unique, across all rows for this event)
  const companionsSet = new Set();
  eventRows.forEach(row => {
    if (row[withCol]) {
      row[withCol].split(',').map(c => c.trim()).filter(c => c).forEach(c => companionsSet.add(c));
    }
  });

  // Build artists list
  const artists = eventRows
    .filter(row => row[artistCol])
    .map(row => ({
      name: row[artistCol],
      role: row[roleCol] || "",
      rating: row[ratingCol] || "",
      notes: row[notesCol] || ""
    }));

  // Sort: headliners first, then others, then support
  artists.sort((a, b) => {
    const aH = a.role.toLowerCase().includes("headline");
    const bH = b.role.toLowerCase().includes("headline");
    const aS = a.role.toLowerCase().includes("support");
    const bS = b.role.toLowerCase().includes("support");
    if (aH && !bH) return -1;
    if (!aH && bH) return 1;
    if (!aS && bS) return -1;
    if (aS && !bS) return 1;
    return 0;
  });

  return {
    date: dateStr,
    venue: venueName,
    city: eventRows[0][cityCol] || "",
    eventName: eventRows[0][eventNameCol] || "",
    companions: [...companionsSet],
    artists: artists,
    success: true
  };
}

function getEventDataForClient(dateStr, venueName) {
  return getEventData(dateStr, venueName);
}

/**
 * Get all events attended with a specific companion
 * @param {string} companionName - Name of the companion
 * @return {Object} - Companion data or error
 */
function getCompanionData(companionName) {
  if (!companionName) return { error: "No companion specified", success: false };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  if (!mainSheet) return { error: "Tracker sheet not found", success: false };

  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];

  const artistCol = headers.indexOf("Artist");
  const dateCol = headers.indexOf("Date");
  const venueCol = headers.indexOf("Venue");
  const cityCol = headers.indexOf("City");
  const roleCol = headers.indexOf("Role");
  const eventNameCol = headers.indexOf("Event Name");
  const ratingCol = headers.indexOf("Rating");
  const withCol = headers.indexOf("Who I Went With");
  const statusCol = headers.indexOf("Status");

  // Find all rows for attended shows that include this companion
  const companionShows = data.slice(1).filter(row => {
    if (row[statusCol] !== "Attended" || !row[withCol]) return false;
    return row[withCol].split(',').map(c => c.trim()).includes(companionName);
  });

  if (companionShows.length === 0) {
    return { companion: companionName, events: [], topArtists: [], topVenues: [], summary: { totalEvents: 0 }, success: true };
  }

  // Group rows into unique events by date + venue
  const eventsMap = {};
  companionShows.forEach(show => {
    let dateStr = "";
    if (show[dateCol] instanceof Date) {
      dateStr = Utilities.formatDate(show[dateCol], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof show[dateCol] === "string") {
      if (show[dateCol].includes('/')) {
        const p = show[dateCol].split('/');
        if (p.length === 3) dateStr = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
      } else {
        dateStr = show[dateCol];
      }
    }

    const venue = show[venueCol] || "";
    const key = `${dateStr}|${venue}`;
    const artist = show[artistCol] || "";

    if (!eventsMap[key]) {
      eventsMap[key] = {
        date: dateStr,
        rawDate: show[dateCol],
        venue: venue,
        city: show[cityCol] || "",
        eventName: show[eventNameCol] || "",
        artists: [],
        roles: {},
        ratings: {}
      };
    }

    if (artist && !eventsMap[key].artists.includes(artist)) {
      eventsMap[key].artists.push(artist);
      eventsMap[key].roles[artist] = show[roleCol] || "";
      eventsMap[key].ratings[artist] = show[ratingCol] || "";
    }
  });

  // Sort events chronologically
  const sortedEvents = Object.values(eventsMap).sort((a, b) => a.date.localeCompare(b.date));

  const events = sortedEvents.map(ev => {
    // Find headliner rating with fallbacks
    let rating = "";
    for (const a of ev.artists) {
      if ((ev.roles[a] || "").toLowerCase().includes("headline") && ev.ratings[a]) {
        rating = ev.ratings[a]; break;
      }
    }
    if (!rating) {
      for (const a of ev.artists) { if (ev.ratings[a]) { rating = ev.ratings[a]; break; } }
    }

    return {
      date: ev.date,
      venue: ev.venue,
      city: ev.city,
      eventName: ev.eventName,
      artists: ev.artists.map(a => ({ name: a, role: ev.roles[a] || "" })),
      rating: rating
    };
  });

  // Time span
  const first = sortedEvents[0];
  const last = sortedEvents[sortedEvents.length - 1];
  let timeSpan = "N/A";
  try {
    if (first.rawDate instanceof Date && last.rawDate instanceof Date) {
      const diffDays = Math.ceil(Math.abs(last.rawDate - first.rawDate) / 86400000);
      const diffYears = (diffDays / 365).toFixed(1);
      if (diffDays === 0) timeSpan = "Same day";
      else if (diffDays < 365) timeSpan = `${diffDays} day${diffDays === 1 ? "" : "s"}`;
      else timeSpan = `${diffYears} year${diffYears === "1.0" ? "" : "s"}`;
    }
  } catch (e) {}

  // Top artists seen together
  const artistCounts = {};
  companionShows.forEach(row => {
    const a = row[artistCol];
    if (a) artistCounts[a] = (artistCounts[a] || 0) + 1;
  });
  const topArtists = Object.entries(artistCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([name, count]) => ({ name, count }));

  // Top venues visited together
  const venueCounts = {};
  events.forEach(ev => {
    if (ev.venue) venueCounts[ev.venue] = (venueCounts[ev.venue] || 0) + 1;
  });
  const topVenues = Object.entries(venueCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([name, count]) => ({ name, count }));

  return {
    companion: companionName,
    events: events,
    topArtists: topArtists,
    topVenues: topVenues,
    summary: {
      totalEvents: events.length,
      totalArtists: companionShows.length,
      uniqueArtists: Object.keys(artistCounts).length,
      firstEvent: first.date,
      lastEvent: last.date,
      timeSpan: timeSpan
    },
    success: true
  };
}

function getCompanionDataForClient(companionName) {
  return getCompanionData(companionName);
}

/**
 * Get all years that have attended events, with counts.
 */
function getYears() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  if (!mainSheet) return { error: "Tracker sheet not found", success: false };

  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  const dateCol = headers.indexOf("Date");
  const statusCol = headers.indexOf("Status");

  const yearCounts = {};
  data.slice(1).forEach(row => {
    if (row[statusCol] !== "Attended") return;
    let year = null;
    if (row[dateCol] instanceof Date) {
      year = row[dateCol].getFullYear();
    } else if (typeof row[dateCol] === "string" && row[dateCol].includes('/')) {
      const p = row[dateCol].split('/');
      if (p.length === 3) year = parseInt(p[2]);
    } else if (typeof row[dateCol] === "string") {
      const d = new Date(row[dateCol]);
      if (!isNaN(d.getTime())) year = d.getFullYear();
    }
    if (year) yearCounts[year] = (yearCounts[year] || 0) + 1;
  });

  // Each row is one artist; convert to unique event counts
  // (reuse event-grouping logic via getYearData summary would be heavier,
  //  so just return raw row counts here — good enough for the list display)
  const years = Object.keys(yearCounts)
    .map(y => ({ year: parseInt(y), count: yearCounts[y] }))
    .sort((a, b) => b.year - a.year);

  return { years, success: true };
}

function getYearsForClient() { return getYears(); }

/**
 * Get full breakdown for a specific year.
 * @param {number} year
 */
function getYearData(year) {
  if (!year) return { error: "No year specified", success: false };
  year = parseInt(year);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  if (!mainSheet) return { error: "Tracker sheet not found", success: false };

  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];

  const artistCol = headers.indexOf("Artist");
  const dateCol = headers.indexOf("Date");
  const venueCol = headers.indexOf("Venue");
  const cityCol = headers.indexOf("City");
  const roleCol = headers.indexOf("Role");
  const eventNameCol = headers.indexOf("Event Name");
  const ratingCol = headers.indexOf("Rating");
  const withCol = headers.indexOf("Who I Went With");
  const statusCol = headers.indexOf("Status");

  const yearShows = data.slice(1).filter(row => {
    if (row[statusCol] !== "Attended") return false;
    let rowYear = null;
    if (row[dateCol] instanceof Date) {
      rowYear = row[dateCol].getFullYear();
    } else if (typeof row[dateCol] === "string") {
      if (row[dateCol].includes('/')) {
        const p = row[dateCol].split('/');
        if (p.length === 3) rowYear = parseInt(p[2]);
      } else {
        const d = new Date(row[dateCol]);
        if (!isNaN(d.getTime())) rowYear = d.getFullYear();
      }
    }
    return rowYear === year;
  });

  if (yearShows.length === 0) {
    return { year, events: [], topArtists: [], topVenues: [], summary: { totalEvents: 0 }, success: true };
  }

  // Group into unique events
  const eventsMap = {};
  yearShows.forEach(show => {
    let dateStr = "";
    if (show[dateCol] instanceof Date) {
      dateStr = Utilities.formatDate(show[dateCol], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof show[dateCol] === "string") {
      if (show[dateCol].includes('/')) {
        const p = show[dateCol].split('/');
        if (p.length === 3) dateStr = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`;
      } else { dateStr = show[dateCol]; }
    }
    const venue = show[venueCol] || "";
    const key = `${dateStr}|${venue}`;
    const artist = show[artistCol] || "";

    if (!eventsMap[key]) {
      eventsMap[key] = {
        date: dateStr, rawDate: show[dateCol],
        venue, city: show[cityCol] || "",
        eventName: show[eventNameCol] || "",
        artists: [], roles: {}, ratings: {},
        companions: new Set()
      };
    }
    if (artist && !eventsMap[key].artists.includes(artist)) {
      eventsMap[key].artists.push(artist);
      eventsMap[key].roles[artist] = show[roleCol] || "";
      eventsMap[key].ratings[artist] = show[ratingCol] || "";
    }
    if (show[withCol]) {
      show[withCol].split(',').map(c => c.trim()).filter(c => c)
        .forEach(c => eventsMap[key].companions.add(c));
    }
  });

  const sortedEvents = Object.values(eventsMap).sort((a, b) => a.date.localeCompare(b.date));

  const events = sortedEvents.map(ev => {
    let rating = "";
    for (const a of ev.artists) {
      if ((ev.roles[a] || "").toLowerCase().includes("headline") && ev.ratings[a]) {
        rating = ev.ratings[a]; break;
      }
    }
    if (!rating) { for (const a of ev.artists) { if (ev.ratings[a]) { rating = ev.ratings[a]; break; } } }

    return {
      date: ev.date,
      venue: ev.venue, city: ev.city, eventName: ev.eventName,
      artists: ev.artists.map(a => ({ name: a, role: ev.roles[a] || "" })),
      companions: [...ev.companions],
      rating
    };
  });

  // Summary
  const artistCounts = {};
  yearShows.forEach(row => {
    const a = row[artistCol];
    if (a) artistCounts[a] = (artistCounts[a] || 0) + 1;
  });
  const venueCounts = {};
  events.forEach(ev => { if (ev.venue) venueCounts[ev.venue] = (venueCounts[ev.venue] || 0) + 1; });

  const ratings = events.map(e => parseFloat(e.rating)).filter(r => !isNaN(r));
  const avgRating = ratings.length > 0
    ? (ratings.reduce((a, b) => a + b, 0) / ratings.length).toFixed(1)
    : "N/A";

  const topArtists = Object.entries(artistCounts).sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count }));
  const topVenues = Object.entries(venueCounts).sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count }));

  return {
    year,
    events,
    topArtists,
    topVenues,
    summary: {
      totalEvents: events.length,
      totalArtists: yearShows.length,
      uniqueArtists: Object.keys(artistCounts).length,
      uniqueVenues: Object.keys(venueCounts).length,
      avgRating
    },
    success: true
  };
}

function getYearDataForClient(year) { return getYearData(year); }

/**
 * Get upcoming and announced shows for the web app.
 */
function getUpcomingShows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Tracker");
  if (!mainSheet) return { error: "Tracker sheet not found", success: false };

  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];

  const artistCol = headers.indexOf("Artist");
  const dateCol = headers.indexOf("Date");
  const venueCol = headers.indexOf("Venue");
  const cityCol = headers.indexOf("City");
  const statusCol = headers.indexOf("Status");
  const ticketVendorCol = headers.indexOf("Ticket Vendor");
  const ticketPriceCol = headers.indexOf("Ticket Price");
  const eventNameCol = headers.indexOf("Event Name");

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const shows = [];
  data.slice(1).forEach(row => {
    const status = row[statusCol];
    if (status !== "Upcoming" && status !== "Announced") return;

    let parsedDate = null;
    if (row[dateCol] instanceof Date) {
      parsedDate = new Date(row[dateCol]);
    } else if (typeof row[dateCol] === "string" && row[dateCol].includes('/')) {
      const p = row[dateCol].split('/');
      if (p.length === 3) parsedDate = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    }
    if (!parsedDate || isNaN(parsedDate.getTime())) return;

    parsedDate.setHours(0, 0, 0, 0);
    const diffDays = Math.ceil((parsedDate - today) / 86400000);
    if (diffDays < 0) return;
    const dateStr = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

    shows.push({
      date: dateStr,
      artist: row[artistCol] || "",
      venue: row[venueCol] || "",
      city: row[cityCol] || "",
      eventName: eventNameCol >= 0 ? (row[eventNameCol] || "") : "",
      status,
      ticketVendor: ticketVendorCol >= 0 ? (row[ticketVendorCol] || "") : "",
      ticketPrice: ticketPriceCol >= 0 ? (row[ticketPriceCol] || "") : "",
      daysUntil: diffDays
    });
  });

  shows.sort((a, b) => a.date.localeCompare(b.date));
  return { shows, success: true };
}

function getUpcomingShowsForClient() { return getUpcomingShows(); }