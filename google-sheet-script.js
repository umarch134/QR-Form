/**
 * ============================================================================
 * GB YOUTH GLOBAL - GOOGLE DRIVE INDIVIDUAL RECORD & SCREENSHOT SYSTEM
 * Connected Account: hassanali15121419@gmail.com
 * ============================================================================
 * 
 * YEH SCRIPT GOOGLE DRIVE MEIN KYA KAREGA?
 * 1. hassanali15121419@gmail.com ki Google Drive mein main folder banayega:
 *    📁 "GB Youth Global - Event Registrations"
 * 
 * 2. Har register hone wale bande ka ALAG RECORD FOLDER banayega:
 *    📁 "Ali Khan - 03451234567 (GBY-123456)"
 *    Aur us bande ke folder ke andar 2 cheezein save hongi:
 *      🖼️ "Payment_Screenshot_Ali_Khan.jpg" (Asli Screenshot Image File)
 *      📄 "Registration_Details_Ali_Khan.txt" (Bande ki mukammal info: Name, Email, Phone, City, Date/Time)
 * 
 * 3. Sath hi Main Folder ke andar ek Master Sheet ("GB Youth Global - All Registrations")
 *    bhi rakhega jiske andar URL ki jagah ASLI SCREENSHOT PHOTO cell ke andar nazar aayegi!
 * 
 * ============================================================================
 * SETUP STEPS (hassanali15121419@gmail.com se connect karne ka tareeqa):
 * ============================================================================
 * 1. Browser mein hassanali15121419@gmail.com se login karein aur yeh link kholein:
 *    https://script.new
 * 2. Wahan pehle se likha chota sa code mita kar yeh poora code paste kar dein.
 * 3. Upar Save (💾 ya Ctrl + S) dabayein.
 * 4. Upar Right side par "Deploy" > "New deployment" par click karein:
 *    - Left side ⚙️ (Select type) par click kar ke "Web app" select karein.
 *    - Description: "GB Youth Drive Record"
 *    - Execute as: "Me (hassanali15121419@gmail.com)"
 *    - Who has access: "Anyone" (Yeh LAZMI "Anyone" select karna hai!)
 * 5. "Deploy" par click karein -> "Authorize access" -> hassanali15121419@gmail.com select karein
 *    -> "Advanced" -> "Go to Untitled project (unsafe)" -> "Allow" par click karein.
 * 6. Ab aapko "Web app URL" (https://script.google.com/macros/s/.../exec) milega.
 *    Woh URL copy kar ke index.html mein GOOGLE_SCRIPT_URL mein paste kar dein!
 * ============================================================================
 */

const DRIVE_OWNER_EMAIL = "hassanali15121419@gmail.com";
const MAIN_DRIVE_FOLDER = "GB Youth Global - Event Registrations";
const RECORD_FILE_NAME = "GB Youth Global - All Registrations";
const SHEET_TAB_NAME = "Registrations";

function doGet() {
  return ContentService
    .createTextOutput("✓ GB Youth Global - Google Drive Record System is ACTIVE (" + DRIVE_OWNER_EMAIL + ")")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    const data = JSON.parse(e.postData.contents);

    const regId = data.regId || ("GBY-" + Math.floor(100000 + Math.random() * 900000));
    const timestamp = data.timestamp || new Date().toLocaleString("en-PK", { timeZone: "Asia/Karachi" });
    const fullName = (data.fullName || "Participant").trim();
    const email = (data.email || "").trim();
    const phone = (data.phone || "").trim();
    const city = (data.city || "N/A").trim();
    const screenshotBase64 = data.screenshotBase64 || "";
    const screenshotMimeType = data.screenshotMimeType || "image/jpeg";

    // Determine image extension
    let ext = ".jpg";
    if (screenshotMimeType.indexOf("png") !== -1) ext = ".png";
    else if (screenshotMimeType.indexOf("webp") !== -1) ext = ".webp";

    // 1. Get or create Main Folder in Google Drive
    const mainFolder = getOrCreateFolder(DriveApp.getRootFolder(), MAIN_DRIVE_FOLDER);

    // 2. Create a SEPARATE Folder for this Participant inside Main Folder
    // Example: "Ali Khan - 03401234567 (GBY-482910)"
    const participantFolderName = fullName + " - " + phone + " (" + regId + ")";
    const participantFolder = mainFolder.createFolder(participantFolderName);

    // 3. Save Participant's Details File inside their individual folder
    const detailsContent =
      "====================================================\n" +
      "   GB YOUTH GLOBAL - PARTICIPANT REGISTRATION\n" +
      "   Event: Stimulation and Career Counseling\n" +
      "====================================================\n\n" +
      "Registration ID  : " + regId + "\n" +
      "Date & Time (PKT): " + timestamp + "\n" +
      "Full Name        : " + fullName + "\n" +
      "Email Address    : " + email + "\n" +
      "Phone / WhatsApp : " + phone + "\n" +
      "City             : " + city + "\n" +
      "Payment Screenshot: Saved in this folder\n" +
      "====================================================\n";

    const cleanName = fullName.replace(/[^a-zA-Z0-9_-]/g, "_");
    participantFolder.createFile("Registration_Details_" + cleanName + ".txt", detailsContent, MimeType.PLAIN_TEXT);

    // 4. Save the ACTUAL Payment Screenshot Image File directly inside the Participant's Folder
    let savedImageFile = null;
    if (screenshotBase64) {
      const decodedBytes = Utilities.base64Decode(screenshotBase64);
      const imageFileName = "Payment_Screenshot_" + cleanName + "_" + regId + ext;
      const imageBlob = Utilities.newBlob(decodedBytes, screenshotMimeType, imageFileName);
      savedImageFile = participantFolder.createFile(imageBlob);
      savedImageFile.setDescription(
        "Full Name: " + fullName + " | Phone: " + phone + " | Email: " + email + " | City: " + city + " | ID: " + regId
      );
    }

    // 5. Also add record + actual embedded image to Master Sheet in Main Folder
    try {
      const ss = getOrCreateSpreadsheetInFolder(mainFolder, RECORD_FILE_NAME);
      let sheet = ss.getSheetByName(SHEET_TAB_NAME);
      if (!sheet) {
        sheet = ss.getSheets()[0];
        sheet.setName(SHEET_TAB_NAME);
      }

      if (sheet.getLastRow() === 0) {
        const headers = [
          "Registration ID",
          "Date & Time (PKT)",
          "Full Name",
          "Email Address",
          "Phone / WhatsApp",
          "City",
          "Payment Screenshot",
          "Verification Status"
        ];
        sheet.appendRow(headers);
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground("#0a1628");
        headerRange.setFontColor("#d4af37");
        headerRange.setFontWeight("bold");
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 130);
        sheet.setColumnWidth(2, 165);
        sheet.setColumnWidth(3, 180);
        sheet.setColumnWidth(4, 210);
        sheet.setColumnWidth(5, 150);
        sheet.setColumnWidth(6, 120);
        sheet.setColumnWidth(7, 220);
        sheet.setColumnWidth(8, 150);
      }

      sheet.appendRow([
        regId,
        timestamp,
        fullName,
        email,
        "'" + phone,
        city,
        "Screenshot Saved in Folder",
        "Pending Verification"
      ]);

      const newRowIndex = sheet.getLastRow();
      sheet.setRowHeight(newRowIndex, 140);

      // Embed the actual screenshot image directly inside Column 7 cell (no URL!)
      if (screenshotBase64) {
        const dataUrl = "data:" + screenshotMimeType + ";base64," + screenshotBase64;
        const cellImage = SpreadsheetApp.newCellImage()
          .setSourceUrl(dataUrl)
          .setAltTextTitle(fullName + " - " + regId)
          .build();
        sheet.getRange(newRowIndex, 7).setValue(cellImage);
      }
    } catch (sheetErr) {
      // Individual folder + screenshot file are already safely saved in Drive
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "success",
        regId: regId,
        participantFolder: participantFolderName
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

function getOrCreateSpreadsheetInFolder(folder, fileName) {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}

  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.openById(file.getId());
  }

  const newSs = SpreadsheetApp.create(fileName);
  const ssFile = DriveApp.getFileById(newSs.getId());
  ssFile.moveTo(folder);
  return newSs;
}
