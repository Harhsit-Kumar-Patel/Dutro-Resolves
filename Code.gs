var FOLDER_ID = "1rQPbR3_Bb2ZdRh4rsTDlQHboqBWRUqIHTE2MjjCJCDFtS3XpXxW5QvAIJ5LSzXT-co-8RTyt";
var SPREADSHEET_ID = "1MnCgjj6SRIufY-xm0ZJcGTblc1ygGvBMC7TizwRXXZQ";
var SHEET_NAME = "Form responses 1";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Prevent race conditions on concurrent row appends

  try {
    var data = JSON.parse(e.postData.contents);
    var uploadedUrls = [];
    var uploadedIds = [];

    // 1. Collect files to process (supports up to 5 screenshots)
    var filesToProcess = [];
    if (Array.isArray(data.files) && data.files.length > 0) {
      var sliced = data.files.slice(0, 5);
      for (var f = 0; f < sliced.length; f++) {
        filesToProcess.push(sliced[f]);
      }
    } else if (data.fileData && data.fileName) {
      // Fallback for single-file payloads
      filesToProcess.push({
        fileData: data.fileData,
        fileName: data.fileName,
        mimeType: data.mimeType || "image/png"
      });
    }

    // 2. Process and save each screenshot to Google Drive
    if (filesToProcess.length > 0) {
      try {
        var folder = DriveApp.getFolderById(FOLDER_ID);
        for (var i = 0; i < filesToProcess.length; i++) {
          var item = filesToProcess[i];
          if (!item.fileData) continue;
          try {
            var decodedBytes = Utilities.base64Decode(item.fileData);
            var blob = Utilities.newBlob(decodedBytes, item.mimeType || "image/png", item.fileName || ("screenshot_" + (i + 1) + ".png"));
            var file = folder.createFile(blob);

            // Make file accessible with link
            try {
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } catch (shareErr) {
              console.warn("Share warning: " + shareErr);
            }

            uploadedUrls.push("https://drive.google.com/open?id=" + file.getId());
            uploadedIds.push(file.getId());
          } catch (fileErr) {
            console.error("Error creating file " + (i + 1) + ": " + fileErr);
          }
        }
      } catch (driveErr) {
        console.error("Drive folder error: " + driveErr);
      }
    }

    var fileUrl = uploadedUrls.length > 0 ? uploadedUrls.join(", ") : "N/A";

    // 3. Locate sheet (works for both container-bound and standalone scripts)
    var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    var targetRow = sheet.getLastRow() + 1; // Capture BEFORE writing

    // 4. Calculate Sequential Ticket ID
    var submissionCount = 0;
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var colAValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var j = 0; j < colAValues.length; j++) {
        var val = colAValues[j][0];
        if (val !== "" && val !== null && val !== undefined) {
          submissionCount++;
        }
      }
    }
    var ticketSeq = submissionCount + 1;
    var ticketId = "#DUT-" + ticketSeq;

    // 5. Ensure header in Column B if not present
    try {
      if (sheet.getRange(1, 2).getValue() === "") {
        sheet.getRange(1, 2).setValue("Ticket ID");
      }
    } catch (hdrErr) {
      console.warn("Header check: " + hdrErr);
    }

    // 6. Write the new row using setValues() + setFormula() instead of appendRow().
    // appendRow() silently copies formulas from the row above — setValues() never does.

    // 6a. Write all plain-text columns (A through J) in one batch call
    sheet.getRange(targetRow, 1, 1, 10).setValues([[
      new Date(),                                 // A: Timestamp
      ticketId,                                   // B: Ticket ID
      data.consultantName || "",                  // C: Consultant Name
      data.projectName || data.companyName || "", // D: Project Name
      data.requestType || "",                     // E: Request Type
      data.urgency || "",                         // F: Urgency Level
      data.title || "",                           // G: Request Title
      data.category || "",                        // H: Category / Module
      data.description || "",                     // I: Description
      fileUrl                                     // J: Screenshot URLs
    ]]);

    // 6b. Write IMAGE formula for each uploaded image, one cell at a time
    for (var k = 0; k < uploadedIds.length; k++) {
      var imageFormula = '=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + uploadedIds[k] + '&sz=w1000", 1), "")';
      sheet.getRange(targetRow, 11 + k).setFormula(imageFormula);
    }

    // 6c. Explicitly blank out all unused image slots through column O (col 15)
    //     This prevents ghost formulas from lingering from any prior row
    var clearStart = 11 + uploadedIds.length;
    var clearEnd = 15; // Column O
    if (clearEnd >= clearStart) {
      sheet.getRange(targetRow, clearStart, 1, clearEnd - clearStart + 1).clearContent();
    }

    SpreadsheetApp.flush();

    // Set row height to 100 so thumbnails and text fit cleanly
    sheet.setRowHeight(targetRow, 100);

    SpreadsheetApp.flush();

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      ticketId: ticketId,
      ticketNumber: ticketSeq,
      totalRequests: ticketSeq,
      count: uploadedUrls.length,
      fileUrls: uploadedUrls,
      fileUrl: fileUrl
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

// GET endpoint: Lets the website fetch the live total count of requests directly from Google Sheets!
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    var totalCount = 0;
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var colAValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < colAValues.length; i++) {
        var val = colAValues[i][0];
        if (val !== "" && val !== null && val !== undefined) {
          totalCount++;
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      totalRequests: totalCount,
      nextTicketId: "#DUT-" + (totalCount + 1)
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// STEP 1 HELPER: Test Drive folder access
function testDriveAccess() {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var file = folder.createFile("test_permission.txt", "Drive access verified!");
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log("SUCCESS! Created file: " + file.getUrl());
}

// STEP 2 HELPER: Fill Column B with #DUT-1, #DUT-2... for all existing rows
function fillTicketIdsForExistingRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();

  sheet.getRange(1, 2).setValue("Ticket ID");

  if (lastRow < 2) {
    Logger.log("No data rows found.");
    return;
  }

  var count = 0;
  for (var r = 2; r <= lastRow; r++) {
    var current = sheet.getRange(r, 2).getValue();
    if (!current || String(current).trim() === "") {
      sheet.getRange(r, 2).setValue("#DUT-" + (r - 1));
      count++;
    }
  }
  Logger.log("Updated " + count + " existing rows with Ticket IDs in Column B.");
}

// STEP 3 HELPER: Clean up ghost images from existing rows and set all row heights to 100
function cleanupGhostImages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 15);

  if (lastRow < 2) {
    Logger.log("No data rows found.");
    return;
  }

  var cleanedRows = 0;
  for (var r = 2; r <= lastRow; r++) {
    // Column J (col 10) has comma-separated URLs
    var fileUrlCell = String(sheet.getRange(r, 10).getValue() || "").trim();
    var expectedImageCount = 0;
    if (fileUrlCell && fileUrlCell !== "N/A") {
      var urls = fileUrlCell.split(",");
      for (var u = 0; u < urls.length; u++) {
        if (urls[u].trim().length > 0) expectedImageCount++;
      }
    }

    // Clear all columns from first unused image slot through column O
    var firstUnusedCol = 11 + expectedImageCount;
    if (lastCol >= firstUnusedCol) {
      sheet.getRange(r, firstUnusedCol, 1, lastCol - firstUnusedCol + 1).clearContent();
      cleanedRows++;
    }

    // Reset row height to 100
    sheet.setRowHeight(r, 100);
  }
  SpreadsheetApp.flush();
  Logger.log("Cleaned ghost images from " + cleanedRows + " rows. All row heights set to 100.");
}

// AUTOMATED TRIGGER HANDLER: Called by the 'On change' trigger.
// Fixes the failure email: "Script function not found: autoResizeNewRows"
function autoResizeNewRows(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.setRowHeight(lastRow, 100);
    }
  } catch (err) {
    console.error("autoResizeNewRows error: " + err);
  }
}
