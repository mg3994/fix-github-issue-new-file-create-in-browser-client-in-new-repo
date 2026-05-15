/**
 * Configuration Constants
 */
const MASTER_FOLDER_ID = 'YOUR_MASTER_FOLDER_ID_HERE'; // Replace with your Drive Folder ID
const USER_SHEET_NAME = 'Users';
const APPT_SHEET_NAME = 'Appointments';

/**
 * 1. Fetches real-time type-ahead suggestions based on user input
 */
function getPlaceSuggestions(inputToken) {
  if (!inputToken || inputToken.length < 3) return []; 
  
  try {
    const response = Maps.newGeocoder().geocode(inputToken);
    if (response.results && response.results.length > 0) {
      return response.results.map(result => result.formatted_address);
    }
    return [];
  } catch (e) {
    console.error("Suggestions processing error: " + e.toString());
    return [];
  }
}

/**
 * 2. Processes a textual address to extract coordinates and compute routing metrics
 */
function processLocationAndMetrics(originLat, originLng, destinationQuery) {
  try {
    const geocode = Maps.newGeocoder().geocode(destinationQuery);
    if (!geocode.results || geocode.results.length === 0) {
      throw new Error("Target destination coordinates could not be resolved.");
    }
    
    const result = geocode.results[0];
    const targetAddress = result.formatted_address;
    const targetLat = result.geometry.location.lat;
    const targetLng = result.geometry.location.lng;

    return calculateMatrixMetrics(originLat, originLng, targetLat, targetLng, targetAddress);
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * 3. Processes raw latitude/longitude from a manual map Pin Drop
 */
function processPinDropMetrics(originLat, originLng, pinLat, pinLng) {
  try {
    const response = Maps.newGeocoder().reverseGeocode(pinLat, pinLng);
    let targetAddress = `Pinned Location (${pinLat.toFixed(4)}, ${pinLng.toFixed(4)})`;
    
    if (response.results && response.results.length > 0) {
      targetAddress = response.results[0].formatted_address;
    }

    return calculateMatrixMetrics(originLat, originLng, pinLat, pinLng, targetAddress);
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * Helper to calculate Driving Distance and Duration Matrix metrics
 */
function calculateMatrixMetrics(originLat, originLng, targetLat, targetLng, targetAddress) {
  const directions = Maps.newDirectionFinder()
    .setOrigin(originLat, originLng)
    .setDestination(`${targetLat},${targetLng}`)
    .setMode(Maps.DirectionFinder.Mode.DRIVING)
    .getDirections();

  let distance = "N/A";
  let duration = "N/A";

  if (directions.routes && directions.routes.length > 0) {
    const route = directions.routes[0].legs[0];
    distance = route.distance.text;
    duration = route.duration.text;
  }

  return {
    status: "success",
    address: targetAddress,
    lat: targetLat,
    lng: targetLng,
    distance: distance,
    duration: duration
  };
}

/**
 * 4. Complete the Appointment Booking & User Folder Management
 */
function createAppointment(bookingData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APPT_SHEET_NAME) || ss.insertSheet(APPT_SHEET_NAME);
    
    const start = new Date(bookingData.dateTimeString);
    const end = new Date(start.getTime() + (60 * 60 * 1000)); 

    // Create Calendar Event
    const event = CalendarApp.getDefaultCalendar().createEvent(
      "Service Appointment: " + bookingData.name, start, end,
      { location: bookingData.address, guests: bookingData.email, sendInvites: true }
    );

    // Synchronize profile and generate secure workspace links
    const registration = registerVerifiedUser(bookingData.email, {
      name: bookingData.name,
      lat: bookingData.lat,
      lng: bookingData.lng,
      address: bookingData.address
    });

    // Write complete telemetry data into Database Sheet
    sheet.appendRow([
      new Date(), 
      bookingData.email, 
      bookingData.name, 
      start, 
      event.getId(), 
      bookingData.address, 
      bookingData.lat,
      bookingData.lng,
      bookingData.currentLat,
      bookingData.currentLng,
      bookingData.distance, 
      bookingData.duration,
      registration.folderId || "Linked Account"
    ]);
    
    return { status: "success" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * 5. User Registration & Drive Directory Provisioning
 */
function registerVerifiedUser(email, userData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET_NAME) || ss.insertSheet(USER_SHEET_NAME);
  
  const data = sheet.getDataRange().getValues();
  const existing = data.find(row => row[1] === email);
  if (existing) return { status: "exists", folderId: existing[3] };

  const parentFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  const userFolder = parentFolder.createFolder("Archive: " + email);
  const folderId = userFolder.getId();
  
  try {
    userFolder.addEditor(email); 
  } catch (e) {
    console.warn("Notice: Scope assignment bypassed for email: " + email);
  }

  sheet.appendRow([new Date(), email, userData.name, folderId, userData.lat, userData.lng, userData.address]);
  
  const welcomeSubject = "Your Account and Storage Directory is Ready";
  const welcomeBody = `Hello ${userData.name},\n\nYour appointment has been logged. Access your dedicated document folder here:\nhttps://drive.google.com/drive/folders/${folderId}`;
  
  try { GmailApp.sendEmail(email, welcomeSubject, welcomeBody); } 
  catch (e) { MailApp.sendEmail(email, welcomeSubject, welcomeBody); }

  return { status: "success", folderId: folderId };
}

/**
 * HTML Web App Entry Point Injection
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Service Portal Command Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
