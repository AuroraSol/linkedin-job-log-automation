/**
 * LinkedIn Job Application Automator
 * This script scans Gmail for LinkedIn application confirmations and 
 * automatically logs the details into a Google Sheet.
 */

function autoLogLinkedInApps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Connects to your specific tab. Change "Search Log" to your tab name.
  const sheet = ss.getSheetByName("Search Log") || ss.getSheets()[0];

  // SEARCH: Look for emails from LinkedIn regarding "application" received in the last 7 days
  // you can update the time frame to your preferences and/or also change the trigger settings
  const threads = GmailApp.search('from:jobs-noreply@linkedin.com "application" newer_than:7d');
  
  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const message of messages) {
      
      // Process only emails that haven't been read yet to avoid duplicates
      if (message.isUnread()) {
        const subject = message.getSubject();
        const body = message.getPlainBody();
        const date = message.getDate();

        let role = "";
        let company = "";

        // 1. EXTRACTION: Pull the Role and Company from the Subject Line
        // This looks for patterns like "application to [Role] at [Company]"
        const match = subject.match(/application.* to (.*) at (.*)/i);
        const sentMatch = subject.match(/application.* sent to (.*)/i);

        if (match) {
          role = match[1].trim();
          company = match[2].trim();
        } else if (sentMatch) {
          company = sentMatch[1].trim();
        }

        // 2. FALLBACK: If the subject is generic, scan the email body for the job details
        // Some of the emails subject line do not contain the role but will search the body intead to grab it
        if (!role || role === "Applied") {
          const lines = body.split('\n');
          for (let i = 0; i < lines.length; i++) {
            if (lines[i].toLowerCase().includes("your application was sent to")) {
              // Grabs the company name from the same line and the role from the next line
              company = company || lines[i].split("to").pop().trim();
              role = lines[i+1]?.trim() || lines[i+2]?.trim() || "Applied";
              break;
            }
          }
        }

        // 3. LINK EXTRACTION: Find the direct LinkedIn job URL in the email body
        // This regex captures long URLs including tracking IDs
        const linkMatch = body.match(/https:\/\/www\.linkedin\.com\/(?:comm\/)?jobs\/view\/[^\s]+/);
        const rawLink = linkMatch ? linkMatch[0] : "";

        // 4. FORMATTING: Create a clickable link for the Google Sheet
        const jobHyperlink = rawLink ? '=HYPERLINK("' + rawLink + '", "' + (company || "View Job") + '")' : "No Link Found";

        // 5. LOGGING: Add the data as a new row in your Work Search Log
        // Order: Date | Method | Company | Job Title | Link (Clickable) | Status
        sheet.appendRow([date, "LinkedIn", company, role, jobHyperlink, "Application Sent"]);
        
        // 6. FINISH: Mark the email as read so the script doesn't log it again
        message.markRead();
      }
    }
  }
}
