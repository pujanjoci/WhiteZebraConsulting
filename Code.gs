// Code.gs - Robust handler for Contact Form
const SHEET_NAME = "Form Responses";

function doPost(e) {
  return handlePostRequest(e);
}

function doGet(e) {
  return handleGetRequest(e);
}

function handlePostRequest(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); // Wait up to 10s for concurrent access

  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return createResponse(500, {
        success: false,
        error: "Sheet '" + SHEET_NAME + "' not found",
      });
    }

    // Parse data: Handle JSON or Form URL Encoded
    let data = {};

    if (e.postData && e.postData.contents) {
      try {
        // Try parsing as JSON first
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        // Not JSON? Maybe it's just raw postData or we rely on e.parameter
        console.log("Not JSON content");
      }
    }

    // If JSON parsing failed or yielded empty, try e.parameter (standard form submission)
    if (Object.keys(data).length === 0 && e.parameter) {
      data = e.parameter;
    }

    // Validate required fields
    if (!data.name || !data.email) {
      return createResponse(400, {
        success: false,
        error: "Name and email are required",
      });
    }

    // Append to sheet
    // Matches sheet columns: Name, Email, Phone, Company, Service, Budget, Message, Newsletter, Timestamp
    // Note: We're adding timestamp at the end if you want, or at start.
    // Let's stick to the previous column structure:
    // name, email, phone, company, service, budget, message, newsletter

    const rowData = [
      data.name,
      data.email,
      data.phone || "",
      data.company || "",
      data.service || "",
      "", // Budget field removed
      data.message || "",
      data.newsletter || "No",
      new Date().toLocaleString("en-US"), // Add timestamp for record keeping
    ];

    sheet.appendRow(rowData);

    // Optional: Send email notification
    sendEmailNotification(data);

    return createResponse(200, {
      success: true,
      message: "Form submitted successfully",
    });
  } catch (error) {
    return createResponse(500, {
      success: false,
      error: error.toString(),
    });
  } finally {
    lock.releaseLock();
  }
}

function handleGetRequest(e) {
  return ContentService.createTextOutput("Service is running").setMimeType(
    ContentService.MimeType.TEXT,
  );
}

function createResponse(code, data) {
  const output = JSON.stringify(data);
  // CORS headers are stripped by GAS usually, but we set JSON mime type
  return ContentService.createTextOutput(output).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function sendEmailNotification(formData) {
  try {
    const recipient = Session.getActiveUser().getEmail();
    const subject = "New Contact Form Submission: " + formData.name;
    const body = `
New contact form submission from Whitezebra Consulting website:

Name: ${formData.name}
Email: ${formData.email}
Phone: ${formData.phone || "Not provided"}
Company: ${formData.company || "Not provided"}
Service Interested In: ${formData.service || "Not specified"}

Newsletter Subscription: ${formData.newsletter || "No"}

Message:
${formData.message || "No message provided"}

Timestamp: ${new Date().toLocaleString("en-US")}
    `;

    MailApp.sendEmail(recipient, subject, body);
  } catch (error) {
    console.error("Error sending email notification:", error);
  }
}
