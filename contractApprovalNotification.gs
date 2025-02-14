/*
 * contractApprovalNotification.gs
 *
 * Description:
 * This Google Apps Script automates the notification process for contract approvals.
 * When a user changes the "Approval Decision" in Column D of the Google Sheet to "Approved",
 * the script (via an installable onEdit trigger) retrieves the client's details from the row:
 *   - Client Full Name from Column B
 *   - Client Email Address from Column E
 *   - Client WhatsApp Number from Column F
 * 
 * It then sends a formatted notification message to a designated Google Chat space using a webhook.
 *
 * This automation resolves the issue of manually tracking contract approvals and ensures that notifications
 * are promptly sent to the relevant team, streamlining the approval process.
 *
 * Setup Instructions:
 * 1. Ensure your Google Sheet is structured as follows:
 *    - Column B: Client Full Name
 *    - Column D: Approval Decision
 *    - Column E: Client Email Address
 *    - Column F: Client WhatsApp Number
 * 2. Update the webhook URL in the sendContractApprovedNotification function with your actual Google Chat webhook.
 * 3. Create an installable onEdit trigger (via the Apps Script editor's Triggers menu) for the onEdit function.
 *
 * Author: [Your Name]
 * Date: [Date]
 */

function onEdit(e) {
  var sheet = e.range.getSheet();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();
  
  // Optionally, check for a specific sheet name. Update "Sheet1" to your actual sheet name.
  if (sheet.getName() !== "Sheet1") return;
  
  // Check if the edit occurred in Column D (Approval Decision) and is not in the header row
  if (editedCol === 4 && editedRow > 1) {
    var approvalStatus = e.value;  // New value in Column D
    if (approvalStatus === "Approved") {
      // Retrieve dynamic values from the edited row
      var clientName = sheet.getRange(editedRow, 2).getValue();      // Column B: Client Full Name
      var clientEmail = sheet.getRange(editedRow, 5).getValue();     // Column E: Client Email Address
      var clientWhatsApp = sheet.getRange(editedRow, 6).getValue();  // Column F: Client WhatsApp Number
      
      // Trigger notification function with dynamic data
      sendContractApprovedNotification(clientName, clientEmail, clientWhatsApp);
    }
  }
}

/**
 * Sends a formatted notification to a Google Chat space via webhook.
 *
 * @param {string} clientName - The client's full name.
 * @param {string} clientEmail - The client's email address.
 * @param {string} clientWhatsApp - The client's WhatsApp number.
 */
function sendContractApprovedNotification(clientName, clientEmail, clientWhatsApp) {
  // Replace with your actual Google Chat webhook URL.
  var webhookUrl = "https://chat.googleapis.com/v1/spaces/AAAA1sc0SAY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=sC9FRu29yy9rLLiK_RtZURqT83sgKsHJYdCyAEu5AzM";
  
  // Construct the message text using Google Chat formatting.
  var messageText = "*Client Contract Approved*\n\n" +
                    "Client: *" + clientName + "*\n" +
                    "Email: *" + clientEmail + "*\n" +
                    "WhatsApp: *" + clientWhatsApp + "*\n" +
                    "Status: *Approved*\n\n" +
                    "A new contract has been generated and is available for review.\n" +
                    "Access the contract here: <https://example.com/contract123|View Contract Document>\n\n" +
                    "For full details, please refer to the <https://docs.google.com/spreadsheets/d/1MNitPwfpXl16w4bRwCOZ9k61ld8fYG7Zo8nRGzqmBJg/edit?usp=sharing|CRM Spreadsheet>.";
  
  // Build the payload for the POST request.
  var payload = {
    "text": messageText
  };
  
  // Configure options for the HTTP request.
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  
  try {
    // Send the POST request to the webhook URL.
    var response = UrlFetchApp.fetch(webhookUrl, options);
    Logger.log("Notification sent. Response: " + response.getContentText());
  } catch (error) {
    Logger.log("Error sending notification: " + error);
  }
}
