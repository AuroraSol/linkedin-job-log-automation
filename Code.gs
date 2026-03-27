/**
 * Unified LinkedIn Job Tracker
 * Automatically logs new applications and updates existing ones to "Declined"
 */
function syncLinkedInApplications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Search Log") || ss.getSheets()[0];
  
  // Search for LinkedIn automated emails from the last 7 days
  const threads = GmailApp.search('from:jobs-noreply@linkedin.com newer_than:7d');

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const message of messages) {
      if (message.isUnread()) {
        const subject = message.getSubject();
        const date = message.getDate();
        let role = "";
        let company = "";
        let status = "Application Sent"; 

        // 1. REJECTION CASE (Subject: "Your application to [Role] at [Company]")
        // We identify these by the lack of "was sent" in the subject
        if (subject.includes("Your application to") && !subject.toLowerCase().includes("was sent")) {
          const rejectMatch = subject.match(/Your application to (.*) at (.*)/i);
          if (rejectMatch) {
            role = rejectMatch[1].trim();
            company = rejectMatch[2].trim();
            status = "Declined - No Interview";
          }
        } 
        // 2. NEW APPLICATION CASE (Subject: "Your application was sent to [Role] at [Company]")
        else if (subject.toLowerCase().includes("was sent") || subject.includes("application was sent to")) {
          const appMatch = subject.match(/application.* to (.*) at (.*)/i);
          const sentMatch = subject.match(/application.* sent to (.*)/i);
          if (appMatch) {
            role = appMatch[1].trim();
            company = appMatch[2].trim();
          } else if (sentMatch) {
            company = sentMatch[1].trim();
          }
        }

        if (company) {
          processEntry(sheet, date, company, role, status);
          message.markRead(); 
        }
      }
    }
  }
}

/**
 * Logic to decide whether to update an existing row or add a new one
 */
function processEntry(sheet, date, company, role, status) {
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  // Search Column C (Company) for a match
  for (let i = 0; i < data.length; i++) {
    const rowCompany = data[i][2] ? data[i][2].toString().trim().toLowerCase() : "";
    if (rowCompany === company.toLowerCase().trim()) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex > -1) {
    // Update existing row Status (Column F)
    sheet.getRange(rowIndex, 6).setValue(status);
    // Fill in Role (Column D) if it's blank
    if (role && !data[rowIndex-1][3]) {
      sheet.getRange(rowIndex, 4).setValue(role);
    }
  } else {
    // Add as a new entry if not found
    sheet.appendRow([date, "LinkedIn", company, role, "", status]);
  }
}
