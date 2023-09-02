# LoanManagementApp

## Description

`LoanManagementApp` is a comprehensive loan management automation tool built on Google App Script. It's designed to work with Google Sheets, providing a user-friendly interface for loan application, approval, and repayment processes. The system offers real-time data processing, email notifications, and monthly repayment reminders, ensuring efficient and smooth loan lifecycle management.

## Features

- **Loan Application Processing**: Automatically generates a unique Loan ID upon application submission and sets the initial loan status.
- **Dynamic Loan Periods**: Sets loan periods based on the requested loan amount.
- **Automated Notifications**: Sends email notifications to applicants upon loan approval or decline, with reasons for decline.
- **Repayment Management**: Calculates outstanding amounts, updates loan statuses, and sends monthly reminders for repayments.
- **Triggers & Automation**: Uses Google App Script triggers to automate various processes, ensuring real-time data processing.

## Setup

1. Open your Google Sheet where you want to implement the `LoanManagementApp`.
2. Go to Extensions > Apps Script.
3. Copy and paste the provided script into the script editor.
4. Save and run the setup functions to initialize the necessary triggers.
5. Ensure you have set the necessary permissions for sending emails and accessing the spreadsheet.

## Usage

1. **Loan Application**: Enter the loan application details in the designated Google Sheet.
2. **Approval/Decline**: Update the 'Approval Status' column to either 'Approved' or 'Declined'. If declined, provide a reason in the 'Comments' column.
3. **Repayment**: Update the 'Repayment Amount' column as repayments are made. The system will automatically calculate the outstanding amount and update the loan status.

## Future Enhancements

- Integration with external databases for more robust data storage.
- Enhanced reporting features for loan statistics and insights.
- Integration with other communication platforms for notifications, such as SMS.

## Contributions

Contributions, bug reports, and feature requests are welcome! Please open an issue or submit a pull request.

## License

This project is open-source and available under the MIT License.

