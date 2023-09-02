function processLoanApplication() {
  var spreadsheetId = "1s2kaz4512zT-eYjuxM3X7mPv4_uHKGXFYf_trjiYw9E"; // Replace with your actual spreadsheet ID
  var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Generate Loan ID using timestamp
  var timestamp = new Date().getTime();
  var loanID = "LOAN-" + timestamp;
  sheet.getRange(lastRow, 3).setValue(loanID); // Assuming Loan ID is in column C
  
  // Set Approval Status as Pending
  sheet.getRange(lastRow, 11).setValue("Pending"); // Assuming Approval Status is in column K
  
  // Set Loan Period based on Loan Amount Requested
  var loanAmount = sheet.getRange(lastRow, 7).getValue(); // Assuming Loan Amount Requested is in column G
  var loanPeriod = loanAmount <= 20000 ? "3 Months" : "6 Months";
  sheet.getRange(lastRow, 12).setValue(loanPeriod); // Assuming Loan Period is in column L
  
  // Repayment Due Date will be set when the loan is approved, so we'll handle that in a separate function
  
  // Outstanding Amount (initially, without repayments)
  var outstandingAmount = loanAmount + (loanAmount * 0.05); // Loan amount + 5% interest
  sheet.getRange(lastRow, 16).setValue(outstandingAmount); // Assuming Outstanding Amount is in column P
}

function approveLoan(row) {
  var spreadsheetId = "1s2kaz4512zT-eYjuxM3X7mPv4_uHKGXFYf_trjiYw9E";
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
  var spreadsheetId = "1s2kaz4512zT-eYjuxM3X7mPv4_uHKGXFYf_trjiYw9E";
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
  var spreadsheetId = "1s2kaz4512zT-eYjuxM3X7mPv4_uHKGXFYf_trjiYw9E"; // Your spreadsheet ID

  // Delete existing triggers (to avoid duplicates)
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

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
  var sheet = SpreadsheetApp.openById("1s2kaz4512zT-eYjuxM3X7mPv4_uHKGXFYf_trjiYw9E").getActiveSheet();
  
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
  var sheet = SpreadsheetApp.openById("1s2kaz4512zT-eYjuxM3X7mPv4_uHKGXFYf_trjiYw9E").getActiveSheet();
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
