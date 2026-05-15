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
 * 2. Processes the final selected location to extract coordinates & route metrics
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

    // Calculate Driving Distance and Duration Matrix metrics
    const directions = Maps.newDirectionFinder()
      .setOrigin(originLat, originLng)
      .setDestination(targetAddress)
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
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * 3. Complete the Appointment Booking & User Folder Management
 */
function createAppointment(bookingData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APPT_SHEET_NAME) || ss.insertSheet(APPT_SHEET_NAME);
    
    const start = new Date(bookingData.dateTimeString);
    const end = new Date(start.getTime() + (60 * 60 * 1000)); // 1 hour duration allocation

    // Create Calendar Event on the hosting account's calendar
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

    // Write record data into Database Sheet
    sheet.appendRow([
      new Date(), 
      bookingData.email, 
      bookingData.name, 
      start, 
      event.getId(), 
      bookingData.address, 
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
 * 4. User Registration & Drive Directory Provisioning
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
    console.warn("Notice: Could not securely assign sharing scopes to email: " + email);
  }

  sheet.appendRow([new Date(), email, userData.name, folderId, userData.lat, userData.lng, userData.address]);
  
  const welcomeSubject = "Your Account and Storage Directory is Ready";
  const welcomeBody = `Hello ${userData.name},\n\nYour appointment has been logged. Access your dedicated document folder here:\nhttps://drive.google.com/drive/folders/${folderId}`;
  
  try { GmailApp.sendEmail(email, welcomeSubject, welcomeBody); } 
  catch (e) { MailApp.sendEmail(email, welcomeSubject, welcomeBody); }

  return { status: "success", folderId: folderId };
}

/**
 * Force Auth helper to ensure scopes are recognized by the user runtime
 */
function forceAuth() {
  CalendarApp.getDefaultCalendar().getName();
}

/**
 * HTML Web App Entry Point Injection
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Service Portal Command Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
