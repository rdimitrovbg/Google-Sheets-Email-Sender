# Google-Sheets-Email-Sender

This Google Apps Script provides functionality to send emails directly from a Google Sheets document.<br>
It includes a custom menu in the Google Sheets interface, allowing users to easily send predefined emails.

## Features
- Custom menu with a "Send Application" option for easy access.
- Ability to customize the recipients and email content.
- Automatic calculation of leave period based on input data in the spreadsheet.
- Option to attach a PDF version of the leave application form.
- Error handling for email sending process.

## Usage
1. Open your Google Sheets document.
2. Navigate to the "Extensions" menu.
3. Click on "📧 Send Application".
4. Choose "📤 Send" to send the predefined email.

## Functions
- `onOpen(e)`: Initializes the custom menu and displays an alert message.
- `createSendMenu()`: Creates the custom menu for sending emails.
- `sendMailAction()`: Initiates the email sending process when the "Send" option is clicked.
- `whoami()`: Retrieves the current user's email address and populates it into the spreadsheet.
- `lock()`, `unlock()`: Placeholder functions for future implementation to lock/unlock the spreadsheet.
- `showConfirmSendEmailAlert()`: Displays a confirmation dialog before sending the email.
- `hideSheets()`: Hides all sheets in the spreadsheet except the specified sheet to be converted to PDF.
- `getDataFromCell(arg_1)`: Retrieves data from a specified cell in the spreadsheet.
- `calendarCalculation()`: Calculates the leave period based on input data in the spreadsheet.
- `sendMail()`: Sends the email with predefined content and attachments.
- `debug()`: Placeholder function for debugging purposes.

## Notes
- Ensure that the necessary data is populated in the spreadsheet before sending emails.
- Customize the `sent_to` array with the desired recipients' email addresses.
- Modify the email subject, body, and attachments as per your requirements.
- Implement additional functionality as needed, such as locking/unlocking the spreadsheet.

## TODO
- TODO :)
