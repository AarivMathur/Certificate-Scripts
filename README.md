# Certificate Generation and Email Automation - README

## Project Overview

This Google Apps Script automates the process of generating certificates for recipients based on form responses stored in a Google Sheet. It creates personalized certificates using a Google Slides template, converts them to PDFs, and emails them to the respective recipients. Additionally, it tracks which recipients have already been sent an email to prevent duplicate emails.

## Workflow Overview

1. **Form Responses Collection:**
   - Responses are stored in a Google Sheet (`Form Responses 1`), with columns capturing:
     - Name
     - Email Address
     - Completion Date
     - Status of whether an email has been sent (`Yes` or `No`).

2. **Google Slides Template:**
   - A Google Slides file is used as a certificate template with placeholders (`{{Name}}`, `{{Date}}`) that will be replaced with actual data.

3. **Email Automation:**
   - After generating the certificates, the script sends them to the recipients via email as a PDF attachment.

## Script Description

### Function: `sendCertificates()`

The script operates as follows:

1. **Access the Spreadsheet:**
   - It retrieves all data from the `Form Responses 1` sheet using `SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");`.

2. **Iterate through Responses:**
   - Loops through each row of data, starting from the second row (since the first row is headers).
   - Extracts the name, email, and completion date for each recipient.
   - Skips any rows that don't have an email or have already been marked as "Email Sent."

3. **Generate Certificate:**
   - For each valid response:
     - Makes a copy of the Google Slides certificate template using the `DriveApp` API.
     - Opens the copy and replaces placeholders (`{{Name}}`, `{{Date}}`) with actual data from the sheet.
   
4. **Save and Send Certificate:**
   - Converts the personalized certificate to a PDF.
   - Sends the certificate as an attachment via email using the `MailApp.sendEmail()` method.
   - Marks the email as "Sent" in the sheet by updating the respective column.

### Key Components:

- **Google Sheet (`Form Responses 1`)**
  - Holds the form responses. Each row represents an individual’s data, and the 5th column (`Email Sent`) tracks whether the certificate has been emailed.
  
- **Google Slides Template**
  - A pre-designed certificate template stored in Google Drive, identified by the file ID `'1QoD35fywHdXRdUfVdU_FTSSGy8UceYnZisUsqGqCRJ4'`. It contains placeholders for dynamic data such as the recipient’s name and the completion date.

- **Email Automation**
  - Automatically sends personalized certificates to the email addresses retrieved from the Google Sheet and logs the status to prevent duplicate emails.

### Example of Email Sent:
```text
Subject: Your Certificate of Completion
Body: Congratulations! Attached is your certificate of completion.
Attachment: [PDF of personalized certificate]
