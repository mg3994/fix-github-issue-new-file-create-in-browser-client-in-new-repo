/**
 * Configuration Constants
 */
const MASTER_FOLDER_ID = 'YOUR_MASTER_FOLDER_ID_HERE'; // Replace with your Drive Folder ID
const USER_SHEET_NAME = 'Users';
const APPT_SHEET_NAME = 'Appointments';

/**
 * 1. Process Location entirely on the Server
 * Converts a text address into Coordinates & calculates Travel Metrics
 * Uses built-in Apps Script Maps Service (No API Key Required)
 */
function processLocationAndMetrics(originLat, originLng, destinationQuery) {
  try {
    // Geocode the text input from the user to find its real address and coordinates
    const geocode = Maps.newGeocoder().geocode(destinationQuery);
    if (!geocode.results || geocode.results.length === 0) {
      throw new Error("Location not found.");
    }
    
    const result = geocode.results[0];
    const targetAddress = result.formatted_address;
    const targetLat = result.geometry.location.lat;
    const targetLng = result.geometry.location.lng;

    // Calculate Driving Distance and Duration
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
 * 2. Complete the Appointment Booking
 */
function createAppointment(bookingData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APPT_SHEET_NAME) || ss.insertSheet(APPT_SHEET_NAME);
    
    const start = new Date(bookingData.dateTimeString);
    const end = new Date(start.getTime() + (60 * 60 * 1000)); // 1 hour duration

    // Create Calendar Event
    const event = CalendarApp.getDefaultCalendar().createEvent(
      "Service: " + bookingData.name, start, end,
      { location: bookingData.address, guests: bookingData.email, sendInvites: true }
    );

    // Register/Verify user and handle Drive folder creation seamlessly
    const registration = registerVerifiedUser(bookingData.email, {
      name: bookingData.name,
      lat: bookingData.lat,
      lng: bookingData.lng,
      address: bookingData.address
    });

    // Log complete data packet to Appointments Sheet
    sheet.appendRow([
      new Date(), 
      bookingData.email, 
      bookingData.name, 
      start, 
      event.getId(), 
      bookingData.address, 
      bookingData.distance, 
      bookingData.duration,
      registration.folderId || "Linked"
    ]);
    
    return { status: "success" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * 3. User Registration & Folder Provisioning
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
    userFolder.addEditor(email); // Assign Writer Permission
  } catch (e) {
    console.warn("Could not add writer permission to: " + email);
  }

  sheet.appendRow([new Date(), email, userData.name, folderId, userData.lat, userData.lng, userData.address]);
  sendUserEmail(email, "Account Configured", "Your secure archive folder is active: https://drive.google.com/drive/folders/" + folderId);

  return { status: "success", folderId: folderId };
}

function sendUserEmail(to, subject, body) {
  try { GmailApp.sendEmail(to, subject, body); } 
  catch (e) { MailApp.sendEmail(to, subject, body); }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Secure Executive Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
