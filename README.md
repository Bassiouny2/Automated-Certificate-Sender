# Automated Certificate Sender for UAE Hackathon

This repository contains a Google Apps Script designed to automate the process of sending personalized participation certificates to participants of the UAE Hackathon. The script reads participant information from a Google Sheets spreadsheet, customizes a PowerPoint certificate template, converts it to a PDF, and sends it as an email attachment.

## Features

- Reads participant data from a Google Sheets spreadsheet
- Customizes a PowerPoint template with participant's name
- Converts the customized PowerPoint to a PDF
- Sends an email with the PDF certificate as an attachment

## Requirements

- Google Apps Script enabled on your Google account
- Google Sheets containing participant data
- A PowerPoint template with placeholders for names
- An HTML email template for personalized emails

## Setup

1. **Google Sheets Setup:**
   - Create a Google Sheets document with columns for participant's first name, last name, email address, and a boolean column to track if the certificate is sent (`isSent`).

2. **PowerPoint Template:**
   - Create a PowerPoint template with placeholders `{{Name}}` for the participant's name.

3. **HTML Email Template:**
   - Create an HTML file (`EmailTemplate.html`) for the email body with a placeholder `{{FName}}` for the participant's first name.

4. **Google Apps Script:**
   - Open your Google Sheets document.
   - Click on `Extensions` > `Apps Script`.
   - Copy and paste the provided script into the Apps Script editor.
   - Update the constants `SUBJECT`, `FROM_NAME`, `HTML_TEMPLATE`, and `PPT_ID` with your details.

## Script Details

The script performs the following steps:
1. Reads data from the Google Sheets document.
2. Loops through each participant's data.
3. If the certificate has not been sent (`isSent` is false):
   - Customizes the PowerPoint template with the participant's name.
   - Converts the customized PowerPoint to a PDF.
   - Sends an email with the PDF attached.
   - Updates the `isSent` column in the spreadsheet.
   - Deletes temporary files created during the process.

## Usage

- To execute the script, open the Apps Script editor and click the run button.
- The script will automatically process and send certificates to participants whose certificates have not yet been sent.

## Example

Here is an example of the data structure expected in the Google Sheets document:

| FirstName | LastName | EmailAddress        | isSent |
|-----------|----------|---------------------|--------|
| John      | Doe      | john.doe@example.com| false  |
| Jane      | Smith    | jane.smith@example.com| false  |

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
