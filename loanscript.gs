function processLoanApplication() {
  var spreadsheetId = "Your google sheet file ID"; // Replace with your actual spreadsheet ID
  var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  var email = sheet.getRange(lastRow, 6).getValue(); // Assuming Email Address is in column F
  
  // Get all the data to check existing loans (excluding the last row)
  var dataRange = sheet.getRange(2, 1, lastRow - 2, 18).getValues();
  
  var existingLoans = dataRange.filter(function(row) {
    return row[5] === email; // Assuming Email Address is in column F (0-indexed)
  });
  
  var totalExistingLoans = existingLoans.reduce(function(sum, loan) {
    return sum + loan[6]; // Assuming Loan Amount Requested is in column G (0-indexed)
  }, 0);

  var maxLoanAmount = 50000;
  var availableLoanAmount = maxLoanAmount - totalExistingLoans;
  
  var hasDefaulted = false;
  var defaultedLoanOutstandingAmount = 0;
  
  existingLoans.forEach(function(loan) {
    if (loan[16] === "Defaulted") { // Assuming Loan Status is in column Q (0-indexed)
      hasDefaulted = true;
      defaultedLoanOutstandingAmount += loan[15]; // Assuming Outstanding Amount is in column P (0-indexed)
    }
  });
  
  var loanAmount = sheet.getRange(lastRow, 7).getValue(); // Assuming Loan Amount Requested is in column G
  
  // Generate Loan ID using timestamp
  var timestamp = new Date().getTime();
  var loanID = "LOAN-" + timestamp;
  sheet.getRange(lastRow, 3).setValue(loanID); // Assuming Loan ID is in column C
  
  // Set Loan Period based on Loan Amount Requested
  var loanPeriod = loanAmount <= 20000 ? "3 Months" : "6 Months";
  sheet.getRange(lastRow, 12).setValue(loanPeriod); // Assuming Loan Period is in column L
  
  // Outstanding Amount (initially, without repayments)
  var outstandingAmount = loanAmount + (loanAmount * 0.05); // Loan amount + 5% interest
  
  if (hasDefaulted) {
    sheet.getRange(lastRow, 11).setValue("Declined"); // Assuming Approval Status is in column K
    var declineReason = "You have a defaulted loan. You need to repay the outstanding amount before applying for another loan. Your outstanding loan amount including interest is: " + defaultedLoanOutstandingAmount;
    sheet.getRange(lastRow, 18).setValue(declineReason); // Assuming Comments is in column R
    outstandingAmount = 0; // Reset outstanding amount to zero
  } else if (loanAmount > availableLoanAmount) {
    sheet.getRange(lastRow, 11).setValue("Declined"); // Assuming Approval Status is in column K
    var declineReason = "The amount requested exceeds your available loan limit. You can apply for a maximum of " + availableLoanAmount + ".";
    sheet.getRange(lastRow, 18).setValue(declineReason); // Assuming Comments is in column R
    outstandingAmount = 0; // Reset outstanding amount to zero
  } else {
    // Set Approval Status as Pending
    sheet.getRange(lastRow, 11).setValue("Pending"); // Assuming Approval Status is in column K
  }
  
  sheet.getRange(lastRow, 16).setValue(outstandingAmount); // Assuming Outstanding Amount is in column P

  // Call the sendEmailNotification function to notify the applicant
  sendEmailNotification(lastRow);
}



function approveLoan(row) {
  var spreadsheetId = "Your google sheet file ID";
  var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  
  var approvalStatus = sheet.getRange(row, 11).getValue(); // K
  
  if (approvalStatus === "Approved") {
    // Set Loan Approval Date
    var approvalDate = new Date();
    sheet.getRange(row, 13).setValue(approvalDate); // M
    
    // Set Repayment Due Date based on Loan Period
    var loanPeriod = sheet.getRange(row, 12).getValue(); // L
    var repaymentDueDate = new Date(approvalDate);
    if (loanPeriod === "3 Months") {
      repaymentDueDate.setMonth(repaymentDueDate.getMonth() + 3);
    } else {
      repaymentDueDate.setMonth(repaymentDueDate.getMonth() + 6);
    }
    sheet.getRange(row, 14).setValue(repaymentDueDate); // N
    
    // Recalculate Outstanding Amount
    var loanAmount = parseFloat(sheet.getRange(row, 7).getValue()); // G
    var outstandingAmount = loanAmount + (loanAmount * 0.05); // Loan amount + 5% interest
    sheet.getRange(row, 16).setValue(outstandingAmount); // P
    
    // Set Loan Status to "In Progress"
    sheet.getRange(row, 17).setValue("In Progress"); // Q
  }
}


function handleRepayment(row) {
  var spreadsheetId = "Your google sheet file ID";
  var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  
  var loanAmount = parseFloat(sheet.getRange(row, 7).getValue()); // G
  var repaymentAmount = parseFloat(sheet.getRange(row, 15).getValue()); // O
  
  // Calculate Outstanding Amount
  var outstandingAmount = (loanAmount + (loanAmount * 0.05)) - repaymentAmount;
  sheet.getRange(row, 16).setValue(outstandingAmount); // P
  
  // Update Loan Status based on Outstanding Amount and Repayment Due Date
  var repaymentDueDate = new Date(sheet.getRange(row, 14).getValue()); // N
  var currentDate = new Date();
  
  if (outstandingAmount <= 0) {
    sheet.getRange(row, 17).setValue("Completed"); // Q
  } else if (currentDate > repaymentDueDate) {
    sheet.getRange(row, 17).setValue("Defaulted"); // Q
  } else {
    sheet.getRange(row, 17).setValue("In Progress"); // Q
  }
}

function setupTriggers() {
  var spreadsheetId = "Your google sheet file ID"; // Your spreadsheet ID

 
  // Create a new onEdit trigger
  ScriptApp.newTrigger('customOnEdit')
    .forSpreadsheet(spreadsheetId)
    .onEdit()
    .create();
}

function customOnEdit(e) {
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  var sheet = e.source.getActiveSheet();
  
  // Check if "Approval Status" column is edited
  if (col === 11) {
    var currentApprovalStatus = sheet.getRange(row, 11).getValue();
    var previousApprovalStatus = e.oldValue;
    
    // If the loan was previously approved, revert any changes to the approval status
    if (previousApprovalStatus === "Approved") {
      range.setValue("Approved");
      Browser.msgBox("A loan that has been approved cannot change its approval status.");
      return;
    }
    
    if (currentApprovalStatus === "Declined" && !sheet.getRange(row, 18).getValue()) {
      sheet.getRange(row, 11).setValue(e.oldValue); // Revert the status
      Browser.msgBox("Please provide a reason in the 'Comments' column when declining a loan application.");
      return;
    }
    
    if (currentApprovalStatus === "Approved" || currentApprovalStatus === "Declined") {
      approveLoan(row);
      sendEmailNotification(row); // Send email notification
    }
  }
  
  // Check if "Repayment Amount" column is edited
  if (col === 15) {
    handleRepayment(row);
  }
  if (currentApprovalStatus === "Declined") {
    sheet.getRange(row, 16).setValue(0); // Assuming Outstanding Amount is in column P
  }
}


function sendEmailNotification(row) {
  var sheet = SpreadsheetApp.openById("Your google sheet file ID").getActiveSheet();
  
  var email = sheet.getRange(row, 6).getValue(); // F
  var approvalStatus = sheet.getRange(row, 11).getValue(); // K
  var fullName = sheet.getRange(row, 4).getValue(); // D
  var loanAmount = sheet.getRange(row, 7).getValue(); // G
  var outstandingAmount = sheet.getRange(row, 16).getValue(); // P
  var comment = sheet.getRange(row, 18).getValue(); // R
  
  var subject = "";
  var body = "";
  
  if (approvalStatus === "Approved") {
    subject = "Loan Application Approved";
    body = "Dear " + fullName + ",\n\nYour loan application has been approved for an amount of " + loanAmount + ". Your repayment amount is " + outstandingAmount + ".\n\nThank you for choosing us!";
  } else if (approvalStatus === "Declined") {
    subject = "Loan Application Declined";
    body = "Dear " + fullName + ",\n\nWe regret to inform you that your loan application has been declined for the following reason: " + comment + "\n\nPlease contact us for more details.";
  }
  
  if (subject && body) {
    try {
      MailApp.sendEmail(email, subject, body);
    } catch (e) {
      Logger.log("Failed to send email to " + email + ". Error: " + e.toString());
    }
  }
}

function sendMonthlyReminders() {
  var sheet = SpreadsheetApp.openById("Your google sheet file ID").getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Assuming the Approval Status is in column K and Loan Status is in column Q
  var approvalStatusRange = sheet.getRange(2, 11, lastRow - 1, 1);
  var loanStatusRange = sheet.getRange(2, 17, lastRow - 1, 1); 
  
  var approvalStatusValues = approvalStatusRange.getValues();
  var loanStatusValues = loanStatusRange.getValues();
  
  for (var i = 0; i < approvalStatusValues.length; i++) {
    var approvalStatus = approvalStatusValues[i][0];
    var loanStatus = loanStatusValues[i][0];
    
    // Check if the approval status is "Approved" and loan status is either "In Progress" or "Defaulted"
    if (approvalStatus === "Approved" && (loanStatus === "In Progress" || loanStatus === "Defaulted")) {
      var email = sheet.getRange(i + 2, 6).getValue(); // Assuming Email Address is in column F
      var outstandingAmount = sheet.getRange(i + 2, 16).getValue(); // Assuming Outstanding Amount is in column P
      var fullName = sheet.getRange(i + 2, 4).getValue(); // Assuming Full Name is in column D
      
      // Ensure the email is valid and the outstanding amount is greater than zero
      if (email && email.includes("@") && outstandingAmount > 0) {
        var subject = "Monthly Loan Repayment Reminder";
        var body = "Dear " + fullName + ",\n\nThis is a reminder about your outstanding loan amount of " + outstandingAmount + ".\n\nPlease ensure timely repayment to avoid any penalties.\n\nThank you for choosing us!";
        
        MailApp.sendEmail(email, subject, body);
      }
    }
  }
}

function setupMonthlyTrigger() {
    
    // Create a new monthly trigger
    ScriptApp.newTrigger('sendMonthlyReminders')
        .timeBased()
        .onMonthDay(1)
        .atHour(9) // You can adjust the hour as per your preference
        .create();
}

function notifyApprover() {
  var spreadsheetId = "Your google sheet file ID"; // Replace with your actual spreadsheet ID
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Application"); // Replace with your actual sheet name
  
  var lastRow = sheet.getLastRow();
  
  var approverEmail = "examples@gmail.com"; // Replace with the actual email address of the approver
  var applicantName = sheet.getRange(lastRow, 4).getValue(); // Assuming Full Name is in column D
  var loanAmount = sheet.getRange(lastRow, 7).getValue(); // Assuming Loan Amount Requested is in column G
  
  var subject = "New Loan Application Submitted";
  var body = "A new loan application has been submitted by " + applicantName + " for an amount of " + loanAmount + ". Please review the application.";
  
  MailApp.sendEmail(approverEmail, subject, body);
}







 
