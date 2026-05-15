/**
 * Configuration
 */
const MASTER_FOLDER_ID = 'YOUR_MASTER_FOLDER_ID_HERE';
const USER_SHEET_NAME = 'Users';

/**
 * Registers a user, creates a folder, and assigns "Writer" permissions.
 * @param {string} email - The verified email from Google/Apple/Manual Auth.
 * @param {Object} userData - Contains name, lat, lng, and address.
 */
function registerVerifiedUser(email, userData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET_NAME) || ss.insertSheet(USER_SHEET_NAME);
    
    // 1. Check for existing user to avoid duplicate folders
    const data = sheet.getDataRange().getValues();
    const existingUser = data.find(row => row[1] === email);
    
    if (existingUser) {
      return { status: "exists", folderId: existingUser[3] };
    }

    // 2. Folder Creation Logic
    const parentFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const userFolder = parentFolder.createFolder("User Archive: " + email);
    const folderId = userFolder.getId();
    
    // 3. ASSIGN WRITER ROLE
    // This gives the user full edit/upload access to this specific folder
    try {
      userFolder.addEditor(email);
    } catch (authError) {
      console.warn("Could not add editor: " + email + ". Ensure it is a valid Google Account.");
      // Note: Apple IDs that aren't linked to a Google account might fail here.
    }

    // 4. Log to Database (Sheet)
    sheet.appendRow([
      new Date(), 
      email, 
      userData.name || 'Anonymous', 
      folderId, 
      userData.lat, 
      userData.lng, 
      userData.address,
      "Writer" // Role status
    ]);

    // 5. Send Welcome & Notification
    const subject = "Access Granted: Your Private Drive Folder";
    const body = "Your registration is complete. You have been granted 'Writer' access to your folder.\n\n" +
                 "Folder Link: https://drive.google.com/drive/folders/" + folderId;
    
    sendUserEmail(email, subject, body);

    return { status: "success", folderId: folderId };

  } catch (e) {
    console.error("Critical Error: " + e.toString());
    return { status: "error", message: "Failed to initialize user environment." };
  }
}

/**
 * Robust Email Delivery
 */
function sendUserEmail(to, subject, body) {
  try {
    // Attempt via Gmail (better formatting/tracking)
    GmailApp.sendEmail(to, subject, body);
  } catch (e) {
    // Fallback to MailApp if Gmail API is restricted
    MailApp.sendEmail(to, subject, body);
  }
}

/**
 * Places/Maps API Distance Logic
 */
function getDistanceMetrics(originLat, originLng, destination) {
  const directions = Maps.newDirectionFinder()
    .setOrigin(originLat, originLng)
    .setDestination(destination)
    .setMode(Maps.DirectionFinder.Mode.DRIVING)
    .getDirections();

  if (directions.routes.length > 0) {
    const route = directions.routes[0].legs[0];
    return {
      distance: route.distance.text,
      duration: route.duration.text
    };
  }
  return { distance: "N/A", duration: "N/A" };
}
