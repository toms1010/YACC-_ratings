// Google Apps Script - YACC 2025 Feedback System - COMPLETE FIXED VERSION

// REQUIRED: This function serves the HTML form
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('YACC 2025 Feedback Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Handle form submissions from frontend
function submitFeedback(data) {
  try {
    console.log("Received feedback data:", JSON.stringify(data));
    
    // Validate data
    if (!data || !data.ratings) {
      throw new Error("No ratings data provided");
    }
    
    // Save to spreadsheet
    const result = saveToSpreadsheet(data);
    
    // Send email notification
    sendNotificationEmail(data, result.rowNumber);
    
    return {
      success: true,
      message: "Thank you! Your feedback has been recorded.",
      rowNumber: result.rowNumber
    };
    
  } catch (error) {
    console.error("Error in submitFeedback:", error);
    throw new Error("Failed to save feedback: " + error.toString());
  }
}

// Save data to Google Sheets
function saveToSpreadsheet(data) {
  const SPREADSHEET_ID = "1aLeO60-DDOmGVsKPgsbd-iITpWr-EgxlwNktPM0JYAs";
  const SHEET_NAME = "Form Responses";
  
  console.log("Opening spreadsheet:", SPREADSHEET_ID);
  
  try {
    // Open spreadsheet
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      console.log("Created new sheet:", SHEET_NAME);
      
      // Add headers
      const headers = [
        "Timestamp",
        "Medical Committee",
        "Technical and Sound Committee",
        "Program Committee",
        "Stage Committee",
        "Food Committee",
        "Accommodation Committee",
        "Registration Committee",
        "Maintenance Committee",
        "Documentation Committee",
        "Testimony & Suggestions"
      ];
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      console.log("Added headers");
    }
    
    // Prepare row data
    const timestamp = data.timestamp ? new Date(data.timestamp) : new Date();
    const ratings = data.ratings || {};
    
    const rowData = [
      timestamp,
      parseInt(ratings.medical) || 0,
      parseInt(ratings.technical) || 0,
      parseInt(ratings.program) || 0,
      parseInt(ratings.stage) || 0,
      parseInt(ratings.food) || 0,
      parseInt(ratings.accommodation) || 0,
      parseInt(ratings.registration) || 0,
      parseInt(ratings.maintenance) || 0,
      parseInt(ratings.documentation) || 0,
      data.testimony || ""
    ];
    
    console.log("Saving row data:", rowData);
    
    // Append to sheet
    sheet.appendRow(rowData);
    const lastRow = sheet.getLastRow();
    
    // Format the row
    sheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // Center align rating columns (columns 2-10)
    for (let i = 2; i <= 10; i++) {
      sheet.getRange(lastRow, i).setHorizontalAlignment("center");
    }
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    
    console.log("Data saved to row:", lastRow);
    
    return {
      rowNumber: lastRow,
      timestamp: timestamp.toISOString()
    };
    
  } catch (error) {
    console.error("Error saving to spreadsheet:", error);
    throw new Error("Failed to save to spreadsheet: " + error.toString());
  }
}

// Send notification email
function sendNotificationEmail(data, rowNumber) {
  try {
    const recipient = "yacc2025connect@gmail.com";
    const subject = "New YACC 2025 Feedback Submission";
    
    // Calculate average rating
    const ratings = Object.values(data.ratings || {});
    const averageRating = ratings.length > 0 
      ? (ratings.reduce((a, b) => a + b, 0) / ratings.length).toFixed(2)
      : "N/A";
    
    // Create email body
    let body = `
New feedback submission received for YACC 2025!

Submission Details:
- Timestamp: ${new Date().toLocaleString()}
- Row Number in Sheet: ${rowNumber}
- Average Rating: ${averageRating}/5

Committee Ratings:
${Object.entries(data.ratings || {}).map(([committee, rating]) => `  • ${committee}: ${rating}/5`).join('\n')}

Testimony/Suggestions:
${data.testimony || "No testimony provided."}

View the complete response in the spreadsheet.

Thank you for using YACC 2025 Feedback System!
    `;
    
    // Send email
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body
    });
    
    console.log("Notification email sent to:", recipient);
    
  } catch (emailError) {
    console.error("Failed to send notification email:", emailError);
    // Don't throw error - email failure shouldn't block form submission
  }
}

// Handle direct POST requests (for form fallback)
function doPost(e) {
  try {
    console.log("Direct POST request received");
    
    let data = {};
    
    // Check if data is in postData (JSON) or parameters (form)
    if (e.postData && e.postData.contents) {
      // JSON data
      data = JSON.parse(e.postData.contents);
    } else if (e.parameters) {
      // Form data
      data = {
        ratings: {
          medical: e.parameters.medical ? parseInt(e.parameters.medical[0]) : 0,
          technical: e.parameters.technical ? parseInt(e.parameters.technical[0]) : 0,
          program: e.parameters.program ? parseInt(e.parameters.program[0]) : 0,
          stage: e.parameters.stage ? parseInt(e.parameters.stage[0]) : 0,
          food: e.parameters.food ? parseInt(e.parameters.food[0]) : 0,
          accommodation: e.parameters.accommodation ? parseInt(e.parameters.accommodation[0]) : 0,
          registration: e.parameters.registration ? parseInt(e.parameters.registration[0]) : 0,
          maintenance: e.parameters.maintenance ? parseInt(e.parameters.maintenance[0]) : 0,
          documentation: e.parameters.documentation ? parseInt(e.parameters.documentation[0]) : 0
        },
        testimony: e.parameters.testimony ? e.parameters.testimony[0] : "",
        timestamp: e.parameters.timestamp ? e.parameters.timestamp[0] : new Date().toISOString()
      };
    }
    
    // Save data
    const result = saveToSpreadsheet(data);
    
    // Return success response
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: "Thank you! Your feedback has been recorded.",
        rowNumber: result.rowNumber
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error("Error in doPost:", error);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: "Error: " + error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test function
function testBackend() {
  console.log("Testing backend functions...");
  
  try {
    // Test spreadsheet access
    const SPREADSHEET_ID = "1aLeO60-DDOmGVsKPgsbd-iITpWr-EgxlwNktPM0JYAs";
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    return {
      success: true,
      message: "Backend is working properly",
      spreadsheet: spreadsheet.getName(),
      sheets: spreadsheet.getSheets().map(s => s.getName())
    };
    
  } catch (error) {
    return {
      success: false,
      message: "Backend test failed: " + error.toString()
    };
  }
}

// Required function
function myFunction() {
  return "YACC 2025 Feedback Form System - Version 2.0";
}
