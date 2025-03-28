function draftEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the subject from cell D2 and the body from cell D5
  var subject = sheet.getRange("D2").getValue();
  var emailBody = sheet.getRange("D5").getValue();
  
  // Get the folder where the current spreadsheet is located
  var currentFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = currentFile.getParents().next();
  
  // Loop through rows starting from row 2 (excluding the header)
  var data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
  
  // Loop through each row (each name and corresponding email addresses)
  data.forEach(function(row) {
    var name = row[0];  // Name in column A
    var emailAddresses = row[1];  // Comma-separated emails in column B
    
    // Skip rows where column A or B is empty
    if (!name || !emailAddresses) {
      return;  // Skip to the next row if there's no name or emails
    }

    // Replace [name] (case-insensitive) in the subject and body with the actual name
    var personalizedSubject = subject.replace(/\[name\]/gi, name);  // Replace [name] in subject
    var personalizedBody = emailBody.replace(/\[name\]/gi, name);  // Replace [name] in body
    
    // Find the file in the folder that matches the name in column A
    var files = parentFolder.getFilesByName(name);
    var fileLink = '';
    
    // If the file with the matching name is found, get its URL and set it to view-only
    if (files.hasNext()) {
      var file = files.next();
      
      // Set the file to "view-only"
      file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      
      // Get the URL of the file
      fileLink = file.getUrl();
      
      // Create a clickable link with HTML format
      var clickableLink = '<a href="' + fileLink + '" target="_blank">Link</a>';
      
      // Replace [link] with the clickable HTML link
      personalizedBody = personalizedBody.replace(/\[link\]/gi, clickableLink);
    } else {
      // If no matching file is found, provide a default message
      personalizedBody = personalizedBody.replace(/\[link\]/gi, 'No file found with the matching name.');
    }

    // Split the email addresses by commas and trim them to ensure no extra spaces
    var emails = emailAddresses.split(",").map(function(email) {
      return email.trim();
    });
    
    // Only create one draft for each name, sending it to all emails
    if (emails.length > 0) {
      var uniqueEmails = Array.from(new Set(emails));  // Remove duplicates in case of duplicate email addresses
      
      // Create one draft with all unique email addresses for each name
      GmailApp.createDraft(uniqueEmails.join(","), personalizedSubject, personalizedBody, {
        htmlBody: personalizedBody  // Set the body as HTML to make the link clickable
      });
    }
  });
}





function createSheetsInFolder() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  
  // Get the values from column A (from A3 onward)
  var names = sheet.getRange('A3:A').getValues();
  
  // Get the current folder of the spreadsheet
  var file = DriveApp.getFileById(spreadsheet.getId());
  var folderIterator = file.getParents();
  
  // Ensure there's at least one folder
  if (!folderIterator.hasNext()) {
    Logger.log("No parent folder found for this spreadsheet.");
    return;
  }

  var folder = folderIterator.next(); // Get the first folder where the spreadsheet is located
  Logger.log("Using folder: " + folder.getName()); // Log the folder name
  
  // Get the "Report Card" sheet from the original spreadsheet
  var reportCardSheet = spreadsheet.getSheetByName("Report Card");
  if (!reportCardSheet) {
    Logger.log("No 'Report Card' sheet found in the source spreadsheet.");
    return;
  }

  // Get the spreadsheet ID dynamically
  var currentSpreadsheetId = spreadsheet.getId();
  
  // Loop through each name in the range
  for (var i = 0; i < names.length; i++) {
    var name = names[i][0];  // Get the name in the row
    
    // Check if the sheet already exists before creating a new one
    if (name && !fileExistsInFolder(name, folder)) {
      Logger.log("Creating sheet for: " + name);  // Log which sheet is being created
      // If the name is not empty and the file does not exist, create a new sheet
      var newFile = SpreadsheetApp.create(name);  // Create new Google Sheet
      
      // Copy the "Report Card" sheet into the new spreadsheet
      reportCardSheet.copyTo(newFile);  // Copy the "Report Card" sheet into the new file

      // Get the default sheet (the one that was created automatically)
      var defaultSheet = newFile.getSheets()[0];
      
      // Delete the default sheet to leave only the copied "Report Card" sheet
      newFile.deleteSheet(defaultSheet);  // Delete the default sheet

      // Optionally, rename the copied sheet (if necessary)
      var copiedSheet = newFile.getSheets()[0];  // Get the copied "Report Card" sheet
      copiedSheet.setName("Report Card");  // Rename it (optional)

      // Construct the dynamic formula for cell B7 with a static range
      var formulaB7 = '=TRANSPOSE(IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + currentSpreadsheetId + '/edit?gid=1460917246","Master!D2:2"))';
      var cellB7 = copiedSheet.getRange("B7");
      cellB7.setFormula(formulaB7);  // Set formula in B7

      // Construct the dynamic formula for cell C7 with the incrementing range (D3:3, D4:4, D5:5, etc.)
      var rowNumberC7 = 3 + i;  // Start at D3 and increment by 1 for each sheet
      var formulaC7 = '=TRANSPOSE(IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + currentSpreadsheetId + '/edit?gid=1460917246","Master!D' + rowNumberC7 + ':' + rowNumberC7 + '"))';
      var cellC7 = copiedSheet.getRange("C7");
      cellC7.setFormula(formulaC7);  // Set formula in C7

      // Construct the dynamic formula for cell G3 with the incrementing range (C3, C4, C5, etc.)
      var rowNumberG3 = 3 + i;  // Start at C3 and increment by 1 for each sheet
      var formulaG3 = '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + currentSpreadsheetId + '/edit?gid=1460917246#gid=1460917246","Master!C' + rowNumberG3 + '")';
      var cellG3 = copiedSheet.getRange("G3");
      cellG3.setFormula(formulaG3);  // Set formula in G3

      // Construct the dynamic formula for cell G12 with the incrementing range (B3, B4, B5, etc.)
      var rowNumberG12 = 3 + i;  // Start at B3 and increment by 1 for each sheet
      var formulaG12 = '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + currentSpreadsheetId + '/edit?gid=1460917246#gid=1460917246","Master!B' + rowNumberG12 + '")';
      var cellG12 = copiedSheet.getRange("G12");
      cellG12.setFormula(formulaG12);  // Set formula in G12

      // Construct the dynamic formula for cell D7 with the incrementing range (D3, D4, D5, etc.)
      var formulaD7 = '=TRANSPOSE(IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + currentSpreadsheetId + '/edit?gid=1460917246","Master!D1:1"))';
      var cellD7 = copiedSheet.getRange("D7");
      cellD7.setFormula(formulaD7);  // Set the constant formula in D7

      // Move the new file to the same folder as the original spreadsheet
      var newFileDrive = DriveApp.getFileById(newFile.getId());  // Get the file object
      folder.addFile(newFileDrive);  // Move the file to the folder
      DriveApp.getRootFolder().removeFile(newFileDrive);  // Remove the file from the root folder
    } else {
      Logger.log("Sheet already exists or name is empty: " + name); // Log when a file exists or name is empty
    }
  }
}

function fileExistsInFolder(fileName, folder) {
  var files = folder.getFilesByName(fileName);
  return files.hasNext();  // If a file with that name exists, return true
}
