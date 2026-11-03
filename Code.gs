var FOLDER_ID = "1rQPbR3_Bb2ZdRh4rsTDlQHboqBWRUqIHTE2MjjCJCDFtS3XpXxW5QvAIJ5LSzXT-co-8RTyt";
var SPREADSHEET_ID = "1MnCgjj6SRIufY-xm0ZJcGTblc1ygGvBMC7TizwRXXZQ";
var SHEET_NAME = "Form responses 1";

// Builds ONE formula text that is identical for every row.
// It uses INDEX($J:$J, ROW()) = "column J of whatever row this cell is in",
// so it can never point at another row, even if Sheets copies/inherits it.
function buildImageFormula(nthImage) {
  var jCell = 'INDEX($J:$J,ROW())';
  return '=IFERROR(IF(OR(' + jCell + '="",' + jCell + '="N/A"),"",'
    + 'IMAGE("https://drive.google.com/thumbnail?id="'
    + '&REGEXEXTRACT(TRIM(INDEX(SPLIT(' + jCell + ',","),1,' + nthImage + ')),"id=([A-Za-z0-9_\\-]+)")'
    + '&"&sz=w1000",1)),"")';
}

function getImageFormulaRow() {
  return [[1, 2, 3, 4, 5].map(buildImageFormula)];
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var data = JSON.parse(e.postData.contents);
    var uploadedUrls = [];

    // 1. Collect files to process (supports up to 5 screenshots)
    var filesToProcess = [];
    if (Array.isArray(data.files) && data.files.length > 0) {
      var sliced = data.files.slice(0, 5);
      for (var f = 0; f < sliced.length; f++) {
        filesToProcess.push(sliced[f]);
      }
    } else if (data.fileData && data.fileName) {
      filesToProcess.push({
        fileData: data.fileData,
        fileName: data.fileName,
        mimeType: data.mimeType || "image/png"
      });
    }

    // 2. Upload each screenshot to Google Drive
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
            try {
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } catch (shareErr) {
              console.warn("Share warning: " + shareErr);
            }
            uploadedUrls.push("https://drive.google.com/open?id=" + file.getId());
          } catch (fileErr) {
            console.error("Error creating file " + (i + 1) + ": " + fileErr);
          }
        }
      } catch (driveErr) {
        console.error("Drive folder error: " + driveErr);
      }
    }

    // Column J: comma-separated Drive URLs (or "N/A" if none)
    var fileUrl = uploadedUrls.length > 0 ? uploadedUrls.join(", ") : "N/A";

    // 3. Locate sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    var targetRow = sheet.getLastRow() + 1;

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

    // 5. Ensure header in Column B
    try {
      if (sheet.getRange(1, 2).getValue() === "") {
        sheet.getRange(1, 2).setValue("Ticket ID");
      }
    } catch (hdrErr) {
      console.warn("Header check: " + hdrErr);
    }

    // 6a. Clear image columns first so nothing inherited/copied can survive
    sheet.getRange(targetRow, 11, 1, 5).clearContent();

    // 6b. Write plain-text columns A through J
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
      fileUrl                                     // J: Comma-separated Drive URLs
    ]]);

    // 6c. Write the row-independent image formulas (K..O = slots 1..5)
    sheet.getRange(targetRow, 11, 1, 5).setFormulas(getImageFormulaRow());

    SpreadsheetApp.flush();
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

// GET endpoint: Returns live total count of requests
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

// STEP 3 HELPER (RUN THIS ONCE): Rewrites K:O for every existing row with the
// row-independent formula, which removes all ghost images currently in the sheet.
function rebuildAllImageFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No data rows found.");
    return;
  }

  var numRows = lastRow - 1;
  var rowFormulas = getImageFormulaRow()[0];
  var allFormulas = [];
  for (var r = 0; r < numRows; r++) {
    allFormulas.push(rowFormulas.slice());
  }

  var range = sheet.getRange(2, 11, numRows, 5);
  range.clearContent();
  range.setFormulas(allFormulas);
  sheet.setRowHeights(2, numRows, 100);

  SpreadsheetApp.flush();
  Logger.log("Rebuilt image formulas for " + numRows + " rows.");
}

// AUTOMATED TRIGGER HANDLER: Called by the 'On change' trigger.
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