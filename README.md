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
### Google Form Creation
1. Navigate to Google Forms.
2. Click on the + Blank form button to create a new form.
3. Name your form, for instance, "Loan Application Form".
5. Add the following fields to your form:
- Full Name: Text field (Required)
- Phone Number: Text field (Required)
- Email Address: Text field (Required)
- Loan Amount Requested: Number field (Required)
- Purpose of the Loan: Dropdown or multiple choice (Options: Business Support, Medical, School Fees, Emergency Expenses, Utility Bills, Other) (Required)
- Do you have an existing Loan: Multiple choice (Options: Yes, No) (Required)
- Terms and Conditions: Paragraph text (Paste the provided "Loan Details" content here)
- I have read and agree to the terms and conditions: Checkbox (Option: I Agree) (Required)
6. Once you've added all the fields, click on the **Send** button on the top right corner. Choose the link option and share the form with your intended audience.

### Generating the Google Sheet
1. After creating the form, click on the Responses tab.
2. Click on the green Sheets icon. This will create a new Google Sheet that's linked to the form. All responses to the form will automatically populate this sheet.

### Script Integration
1. Open the generated Google Sheet.
2. Go to Extensions > Apps Script.
3. Copy and paste the provided LoanManagementApp script into the script editor.
4. Save and run the setup functions to initialize the necessary triggers.
5. Ensure you have set the necessary permissions for sending emails and accessing the spreadsheet.

## Usage

1. **Loan Application**: Users fill out the Google Form for loan applications..
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

