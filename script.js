// Code.gs (Google Apps Script)

const COLUMN_ENTRY_NUMBER = 1;
const COLUMN_DEPARTURE = 5;
const COLUMN_ARRIVAL = 6;

function doGet(e) {
  const sheet = getSheetByName('PP_LogSheet'); // Get sheet here
  if (e && e.parameter && e.parameter.page === 'guard') {
    return HtmlService.createTemplateFromFile('guard').evaluate();
  } else {
    return HtmlService.createTemplateFromFile('index').evaluate();
  }
}

function getSheetByName(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`Sheet '${name}' not found or accessible.`);
  }
  return sheet;
}

function validateEntryNumbers() {
  const sheet = getSheetByName('PP_LogSheet');
  const entryNumbers = sheet.getRange('A2:A').getValues().flat();
  entryNumbers.forEach((entry, index) => {
    if (!entry || !/^\d{8}-\d{4}$/.test(entry)) {
      Logger.log(`Invalid entry at row ${index + 2}: ${entry}`);
    }
  });
}

function generateEntryNumber() {
  const logSheet = getSheetByName('PP_LogSheet');
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');

  let entryNumber;
  let isDuplicate = true;
  let tryCount = 0;

  do {
    // Find the first empty row in column A (Entry Number column):
    let lastRow = 2;  // Start from row 2 (assuming row 1 is for headers)
    while (logSheet.getRange(lastRow, 1).getValue() !== "") {
      lastRow++;
    }


    if (lastRow === 2) { // Sheet is empty except for header
        entryNumber = `${year}${month}${day}-0001`;

    }else {

         const lastEntryNumber = logSheet.getRange(lastRow - 1, 1).getValue(); //Get value from above


          const lastSequence = parseInt(lastEntryNumber.split('-')[1], 10);
            const newSequence = lastSequence + 1;
           entryNumber = `${year}${month}${day}-${String(newSequence).padStart(4, '0')}`;
        }


    isDuplicate = isEntryNumberDuplicate(entryNumber);
    tryCount++;
  } while (isDuplicate && tryCount < 1000);


  if (isDuplicate) {
    Logger.log('Error: Could not generate a unique entry number after multiple tries.');
    return null; 
  }

  return entryNumber;
}

// Function to check if entryNumber already exists in the 'PP_LogSheet'
function isEntryNumberDuplicate(entryNumber) {
  const logSheet = getSheetByName('PP_LogSheet');
  const entryNumbers = logSheet.getRange('A2:A').getValues().flat(); //entryNumbers in column A
  return entryNumbers.includes(entryNumber);
}
// Function to check if entryNumber already exists in the 'PP_LogSheet'
function isEntryNumberDuplicate(entryNumber) {
  const logSheet = getSheetByName('PP_LogSheet');
  const entryNumbers = logSheet.getRange('A2:A').getValues().flat(); //entryNumbers in column A
  return entryNumbers.includes(entryNumber);
}

function submitPass(fullName, typeOfPass, purpose, unit, employeeId, email, date, obDetails, obPurposes, combinedOBDetails) {
  try {
    const logSheet = getSheetByName('PP_LogSheet');
    const personnelPass = getSheetByName('PP_Template');
    const atrbSheet = getSheetByName('ATRB');
    const destinationSheet = getSheetByName("Destination");
    const routingSlip = getSheetByName('ROUTING_SLIP');

    const entryNumber = generateEntryNumber();
    const timestamp = new Date();
    
    // Append data to LogSheet
    logSheet.appendRow([
      entryNumber,
      timestamp,
      fullName,
      typeOfPass,
      "",
      "",
      date,
      "",
      purpose,
      unit,
      employeeId,
      email,
    ]);

    // Populate template sheet
    personnelPass.getRange('I8').setValue(date);
    personnelPass.getRange('F12').setValue(fullName);
    personnelPass.getRange('C19').setValue(purpose);

    // Populate the ROUTING_SLIP
    routingSlip.getRange('E2').setValue(unit);
    routingSlip.getRange('L2').setValue(unit);

    // Handle OB-specific logic
    if (typeOfPass === "OB") {
    personnelPass.getRange('B13').setValue(true);
    personnelPass.getRange('D13').setValue(false);

    // Populate the ATRB sheet
    atrbSheet.getRange('I10').setValue(date);
    atrbSheet.getRange('B18').setValue(date);
    atrbSheet.getRange('C12').setValue(fullName);

    // Populate ATRB: Destination in Column 1, Specific Purpose in Column 4
    let startRow = 21; // Adjust based on where data starts
    for (let i = 0; i < obDetails.length; i++) {
      if (!obDetails[i] || !obPurposes[i]) continue; // Skip if any data is missing
      atrbSheet.getRange(startRow + i, 1).setValue(obDetails[i]);  // Destination
      atrbSheet.getRange(startRow + i, 4).setValue(obPurposes[i]); // Purpose
  }

  // Populate the Destination sheet
  const destinationRow = [
    entryNumber, // Entry Number
    fullName,    // Full Name
  ];

  combinedOBDetails.forEach((detail) => {
    destinationRow.push(detail); // Add combined "Destination:Purpose"
  });

  destinationSheet.appendRow(destinationRow);
}

    else if (typeOfPass === "PB") {
      personnelPass.getRange('B13').setValue(false);
      personnelPass.getRange('D13').setValue(true);

      // Clear ATRB sheet cells if PB
      atrbSheet.getRange('I10').clearContent();
      atrbSheet.getRange('C12').clearContent();
    }

    // Pass the entry number and typeOfPass to exportAndSend
    exportAndSend(email, entryNumber, typeOfPass);

    return entryNumber;
  } catch (error) {
    Logger.log('Error in submitPass: ' + error.message);
    return null;
  }
}

function checkEntryExists(entryNumber) {
  try {
    const sheet = getSheetByName('PP_LogSheet');
    const entryNumbers = sheet.getRange(2, COLUMN_ENTRY_NUMBER, sheet.getLastRow() - 1).getValues().flat(); // Assuming data starts from row 2

    // Check if the entry number exists in the sheet
    const entryExists = entryNumbers.includes(entryNumber);
    return entryExists; // Return true if exists, false otherwise
  } catch (error) {
    Logger.log(`Error in checkEntryExists(${entryNumber}): ${error}`);
    return false;
  }
}

function logTime(action, entryNumber) { // Add entryNumber as a parameter
  try {
    const sheet = getSheetByName('PP_LogSheet');
    const entryNumbers = sheet.getRange(2, COLUMN_ENTRY_NUMBER, sheet.getLastRow() - 1).getValues().flat(); // Assuming data starts from row 2
    const rowIndex = entryNumbers.indexOf(entryNumber);

    if (rowIndex === -1) {
      throw new Error(`Entry number ${entryNumber} not found in the log sheet.`);
    }
   
    const row = rowIndex + 2; // Add 2 to account for header row and array index starting from 0
    const timeStamp = new Date();
    const column = (action === 'departure') ? COLUMN_DEPARTURE : COLUMN_ARRIVAL;

    sheet.getRange(row, column).setValue(timeStamp);


  } catch (error) {
    Logger.log(`Error in logTime(${action}, ${entryNumber}): ${error}`); 
  }
}

function updateTimestamp(type, entryNumber) {
  try {
    const sheet = getSheetByName('PP_LogSheet');
    const data = sheet.getRange(1, COLUMN_ENTRY_NUMBER, sheet.getLastRow()).getValues(); // Get all entry numbers
    const entryNumberExists = data.flat().includes(entryNumber); // Check if entry number exists

     if (!entryNumberExists) {
            // Log the error for debugging in Apps Script
            Logger.log(`Entry Number '${entryNumber}' not found.`); 
            return "Error: Entry Number not found"; // Return an error message
        }

    const row = data.findIndex(row => row[0] === entryNumber) + 1; // Find the row index
    const column = type === 'departure' ? COLUMN_DEPARTURE : COLUMN_ARRIVAL;
    sheet.getRange(row, column).setValue(new Date());
  } catch (error) {
        Logger.log(`Error in updateTimestamp(${type}, ${entryNumber}): ${error}`);
        return error.message; // Or a more user-friendly message
    }
}

function exportAndSend(userEmail, entryNumber, typeOfPass) {
  try {
    if (!userEmail || !validateEmail(userEmail)) {
      throw new Error("Invalid email address provided.");
    }

    const spreadsheetId = '12-tEacZujGE_tBH-tOXQyndp20Hpt5eTCliyu5fi0YU';
    const folderId = '13IVwNZpxNAJEAOD7b6vTOD574TKnPAr5';
    const templateSheetName = 'PP_Template';
    const atrbSheetName = 'ATRB';
    const routingSlipSheetName = 'ROUTING_SLIP';

    const originalSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const templateSheet = originalSpreadsheet.getSheetByName(templateSheetName);
    const atrbSheet = originalSpreadsheet.getSheetByName(atrbSheetName);

    // Create a temporary spreadsheet
    const tempSpreadsheet = SpreadsheetApp.create(`Copy for: ${entryNumber}`);
    const tempSheet = tempSpreadsheet.getSheets()[0];

    // Copy template sheet to the temporary spreadsheet
    const copiedTemplateSheet = templateSheet.copyTo(tempSpreadsheet);
    tempSpreadsheet.deleteSheet(tempSheet); // Remove default sheet
    copiedTemplateSheet.setName(templateSheetName);

    // For OB pass, also copy ATRB sheet
    if (typeOfPass === "OB") {
      const copiedAtrbSheet = atrbSheet.copyTo(tempSpreadsheet).setName(atrbSheetName);
    }

    // Export PP_Template and ATRB as PDF
    const pdfBlobPP = exportSpreadsheetToPdf(tempSpreadsheet.getId());
    const pdfNamePP = `PP_Template-${entryNumber}.pdf`;

    // Export Routing Slip separately
    const routingSlipBlob = exportSheetToPdf(originalSpreadsheet.getId(), routingSlipSheetName, true);
    const routingSlipName = `RoutingSlip-${entryNumber}.pdf`;

    // Save files to Google Drive
    const folder = DriveApp.getFolderById(folderId);
    folder.createFile(pdfBlobPP).setName(pdfNamePP);
    folder.createFile(routingSlipBlob).setName(routingSlipName);

    // Send email with both PDFs
    sendEmail(entryNumber, userEmail, [
      pdfBlobPP.setName(pdfNamePP),
      routingSlipBlob.setName(routingSlipName),
    ]);

    Logger.log(`Personnel Pass submitted successfully! Entry Number: ${entryNumber}`);

    // Clear ATRB and PP_Template after export
    [atrbSheet.getRange('A21:A30'), atrbSheet.getRange('D21:D30')].forEach(range => range.clearContent());


    templateSheet.getRange('I8').clearContent();
    templateSheet.getRange('F12').clearContent();
    templateSheet.getRange('B13').clearContent();
    templateSheet.getRange('D13').clearContent();
    templateSheet.getRange('C19').clearContent();

    Logger.log('ATRB and PP_Template have been cleared for new entries.');

  } catch (error) {
    Logger.log(`Error in exportAndSend: ${error.message}`);
  }
}


// Helper to export entire spreadsheet as PDF
function exportSpreadsheetToPdf(spreadsheetId) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&size=A4&portrait=true&fitw=true&sheetnames=true&printtitle=true&pagenumbers=true`;
  const token = ScriptApp.getOAuthToken();
  const options = {
    headers: {
      Authorization: `Bearer ${token}` // Use template literal here
    }
  };
  return UrlFetchApp.fetch(url, options).getBlob();
}

// Helper to export a specific sheet to PDF with optional landscape orientation
function exportSheetToPdf(spreadsheetId, sheetName, landscape = false) {
  const sheetId = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName).getSheetId();
  const orientation = landscape ? 'landscape' : 'portrait';
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&size=A4&portrait=${!landscape}&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gid=${sheetId}`;
  const token = ScriptApp.getOAuthToken();
  const options = {
    headers: {
      Authorization: `Bearer ${token}`  
    }
  };
  return UrlFetchApp.fetch(url, options).getBlob();
}

// Helper function to validate email
function validateEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}


function sendEmail(entryNumber, email, pdfBlobs) {
  const subject = "Your Personnel Pass Request";
  const body = `Here's your requested Personnel Pass Entry Number: ${entryNumber}`;

  if (!Array.isArray(pdfBlobs)) {
    throw new Error('Expected an array of PDF blobs for attachments.');
  }

  // Ensure all blobs are PDFs
    pdfBlobs.forEach((blob, index) => {
      if (blob.getContentType() !== 'application/pdf') {
          throw new Error(`Attachment at index ${index} is not a valid PDF blob.`);
      }
  });

  // Send the email
  GmailApp.sendEmail(email, subject, body, {
    attachments: pdfBlobs,
    replyTo: 'no-reply@adminsciptbybibohthings.psa',
  });
}


//Powered By © BibohThings. This Personnel Pass System is provided only for PSA Palawan and not intended for sale.
