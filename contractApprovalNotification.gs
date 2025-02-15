/*
 * contractApprovalNotification.gs
 *
 * Description:
 * This script monitors a Google Sheet (e.g., "Sheet1") for changes in the "Approval Decision" column (Column D).
 * When the status is changed to "Approved", the script retrieves:
 *   - Client Full Name from Column B
 *   - Client Email Address from Column E
 *   - Client Phone (WhatsApp) from Column F
 *   - Service Type from Column I
 *   - Project Description from Column Q
 *   - Preferred Service Delivery Date from Column L
 *   - Contract Date as the date the Approval status was changed
 *
 * It then sends two webhooks:
 *   1. A Google Chat webhook to notify the team.
 *   2. A Make.com webhook to trigger document creation from a template.
 *
 * This automation streamlines the contract approval process and ensures full data flexibility for mapping.
 *
 * Setup:
 * 1. Confirm your Google Sheet structure:
 *    - Column B: Client Full Name
 *    - Column D: Approval Decision
 *    - Column E: Client Email Address
 *    - Column F: Client Phone/WhatsApp Number
 *    - Column I: Service Type
 *    - Column Q: Project Description
 *    - Column L: Preferred Service Delivery Date
 * 2. Update the sheet name ("Sheet1") if necessary.
 * 3. Replace the webhook URLs with your actual URLs.
 * 4. Create an installable onEdit trigger for the onEdit function.
 *
 * Author: [Your Name]
 * Date: [Date]
 */

function onEdit(e) {
  var sheet = e.range.getSheet();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();
  
  // Only run on the specified sheet (update "Sheet1" if needed)
  if (sheet.getName() !== "Sheet1") return;
  
  // Proceed if the edited cell is in Column D (Approval Decision) and not in the header row
  if (editedCol === 4 && editedRow > 1) {
    var approvalStatus = e.value; // New value in Column D
    if (approvalStatus === "Approved") {
      // Retrieve key client details and additional data:
      var clientName = sheet.getRange(editedRow, 2).getValue();       // Column B: Client Full Name
      var clientEmail = sheet.getRange(editedRow, 5).getValue();        // Column E: Client Email Address
      var clientPhone = sheet.getRange(editedRow, 6).getValue();        // Column F: Client Phone/WhatsApp Number
      var serviceType = sheet.getRange(editedRow, 9).getValue();        // Column I: Service Type
      var projectDetails = sheet.getRange(editedRow, 17).getValue();    // Column Q: Project Description
      var preferredServiceDeliveryDate = sheet.getRange(editedRow, 12).getValue(); // Column L: Preferred Service Delivery Date
      
      // Capture the contract date as the current date when approval is set
      var contractDate = new Date();
      
      // Retrieve the entire row for additional flexibility if needed:
      var rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // Trigger both webhooks:
      sendGoogleChatNotification(clientName, clientEmail, clientPhone, contractDate);
      triggerMakeWebhook(clientName, clientEmail, clientPhone, serviceType, projectDetails, preferredServiceDeliveryDate, contractDate, rowData);
    }
  }
}

/**
 * Sends a notification to Google Chat via webhook.
 *
 * @param {string} clientName - Client's full name.
 * @param {string} clientEmail - Client's email address.
 * @param {string} clientPhone - Client's phone/WhatsApp number.
 * @param {Date} contractDate - The date the approval was made.
 */
function sendGoogleChatNotification(clientName, clientEmail, clientPhone, contractDate) {
  var googleChatWebhookUrl = "https://chat.googleapis.com/v1/spaces/AAAA1sc0SAY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=sC9FRu29yy9rLLiK_RtZURqT83sgKsHJYdCyAEu5AzM";
  
  var messageText = "*Client Contract Approved*\n\n" +
                    "Client: *" + clientName + "*\n" +
                    "Email: *" + clientEmail + "*\n" +
                    "Phone: *" + clientPhone + "*\n" +
                    "Contract Date: *" + contractDate.toLocaleDateString() + "*\n" +
                    "Status: *Approved*\n\n" +
                    "A new contract has been generated and is available for review.\n" +
                    "Access the contract folder here: <https://drive.google.com/drive/u/0/folders/1wf_4Ajhgpnbwpzo5-D_WAfBKxELh2EKq|View Contract Folder>\n\n" +
                    "For full details, please refer to the <https://docs.google.com/spreadsheets/d/1MNitPwfpXl16w4bRwCOZ9k61ld8fYG7Zo8nRGzqmBJg/edit?usp=sharing|CRM Spreadsheet>.";
  
  var payload = { "text": messageText };
  
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  
  try {
    var response = UrlFetchApp.fetch(googleChatWebhookUrl, options);
    Logger.log("Google Chat notification sent. Response: " + response.getContentText());
  } catch (error) {
    Logger.log("Error sending Google Chat notification: " + error);
  }
}

/**
 * Triggers the Make.com webhook to start the document creation process.
 *
 * @param {string} clientName - Client's full name.
 * @param {string} clientEmail - Client's email address.
 * @param {string} clientPhone - Client's phone/WhatsApp number.
 * @param {string} serviceType - Service Type from Column I.
 * @param {string} projectDetails - Project Description from Column Q.
 * @param {string} preferredServiceDeliveryDate - Preferred Service Delivery Date from Column L.
 * @param {Date} contractDate - The date the contract was approved.
 * @param {Array} rowData - Entire row data for additional mapping flexibility.
 */
function triggerMakeWebhook(clientName, clientEmail, clientPhone, serviceType, projectDetails, preferredServiceDeliveryDate, contractDate, rowData) {
  var makeWebhookUrl = "https://hook.us1.make.com/wkyq6qhl8hweawgrduw6vx83t1a1oz3g";
  
  // Build a comprehensive payload for Make.com:
  var payload = {
    clientName: clientName,
    clientEmail: clientEmail,
    clientPhone: clientPhone,
    approvalStatus: "Approved",
    serviceType: serviceType,
    projectDetails: projectDetails,
    preferredServiceDeliveryDate: preferredServiceDeliveryDate,
    contractDate: contractDate,
    rowData: rowData
  };
  
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  
  try {
    var response = UrlFetchApp.fetch(makeWebhookUrl, options);
    Logger.log("Make.com webhook triggered. Response: " + response.getContentText());
  } catch (error) {
    Logger.log("Error triggering Make.com webhook: " + error);
  }
}
