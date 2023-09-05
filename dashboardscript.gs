function updateDashboard() {
  var spreadsheetId = "Your google sheet file ID";
  var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  var dashboard = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Dashboard");
  
  var lastRow = sheet.getLastRow();
  
  // Get all the data
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
  
  var totalLoans = 0;
  var totalAmountLoaned = 0;
  var defaultedLoans = 0;
  var loanAmountBins = { '0-5000': 0, '5001-10000': 0, '10001-15000': 0, '15001-20000': 0, '20001+': 0 };
  var loanPurposes = {};
  var loanStatuses = { 'Approved': 0, 'Declined': 0, 'In Progress': 0, 'Defaulted': 0, 'Completed': 0 };
  var loanPeriods = { '3 Months': 0, '6 Months': 0 };
  var outstandingAmount = 0;
  var repaymentTimes = { 'On Time': 0, '1 Week Late': 0, '2 Weeks Late': 0, '3+ Weeks Late': 0 };
  var comments = [];
  
  for (var i = 0; i < dataRange.length; i++) {
    var row = dataRange[i];
    var approvalStatus = row[10];
    var loanAmount = row[6];
    var loanPurpose = row[8];
    var loanPeriod = row[11];
    var repaymentDate = row[12];
    var dueDate = row[13];
    var comment = row[14];
    
    // Total Loans and Amount
    if (approvalStatus === "Approved") {
      totalLoans++;
      totalAmountLoaned += loanAmount;
    }
    
    // Defaulted Loans
    if (approvalStatus === "Defaulted") {
      defaultedLoans++;
    }
    
    // Loan Amount Distribution
    if (loanAmount <= 5000) {
      loanAmountBins['0-5000']++;
    } else if (loanAmount <= 10000) {
      loanAmountBins['5001-10000']++;
    } else if (loanAmount <= 15000) {
      loanAmountBins['10001-15000']++;
    } else if (loanAmount <= 20000) {
      loanAmountBins['15001-20000']++;
    } else {
      loanAmountBins['20001+']++;
    }
    
    // Loan Purpose Distribution
    if (!loanPurposes[loanPurpose]) {
      loanPurposes[loanPurpose] = 0;
    }
    loanPurposes[loanPurpose]++;
    
    // Loan Status Distribution
    if (loanStatuses[approvalStatus] !== undefined) {
      loanStatuses[approvalStatus]++;
    }
    
    // Loan Period Distribution
    if (loanPeriods[loanPeriod] !== undefined) {
      loanPeriods[loanPeriod]++;
    }
    
    // Outstanding Amount
    if (approvalStatus === "In Progress" || approvalStatus === "Defaulted") {
      outstandingAmount += loanAmount; 
    }
    
    // Average Repayment Time
    if (repaymentDate && dueDate) {
      var daysLate = (new Date(repaymentDate) - new Date(dueDate)) / (1000 * 60 * 60 * 24);
      if (daysLate <= 0) {
        repaymentTimes['On Time']++;
      } else if (daysLate <= 7) {
        repaymentTimes['1 Week Late']++;
      } else if (daysLate <= 14) {
        repaymentTimes['2 Weeks Late']++;
      } else {
        repaymentTimes['3+ Weeks Late']++;
      }
    }
    
    // Comments for Decline
    if (approvalStatus === "Declined" && comment) {
      comments.push(comment);
    }
  }
  
  // Update the Dashboard sheet with the calculated metrics

  dashboard.getRange("A1").setValue("Total Loans");
  dashboard.getRange("A2").setValue(totalLoans);
  
  dashboard.getRange("B1").setValue("Total Amount Loaned");
  dashboard.getRange("B2").setValue(totalAmountLoaned);
  
  dashboard.getRange("C1").setValue("Defaulted Loans");
  dashboard.getRange("C2").setValue(defaultedLoans);
  
 // Loan Amount Distribution
  var colIndex = 4; // Starting from column D for this example
  for (var bin in loanAmountBins) {
    dashboard.getRange(1, colIndex).setValue(bin);
    dashboard.getRange(2, colIndex).setValue(loanAmountBins[bin]);
    colIndex++;
  }

  // Loan Purpose Distribution
  var loanPurposeIndex = colIndex;
  for (var purpose in loanPurposes) {
    dashboard.getRange(1, loanPurposeIndex).setValue(purpose);
    dashboard.getRange(2, loanPurposeIndex).setValue(loanPurposes[purpose]);
    loanPurposeIndex++;
  }

  // Loan Status Distribution
  var loanStatusIndex = loanPurposeIndex;
  for (var status in loanStatuses) {
    dashboard.getRange(1, loanStatusIndex).setValue(status);
    dashboard.getRange(2, loanStatusIndex).setValue(loanStatuses[status]);
    loanStatusIndex++;
  }

  // Loan Period Distribution
  var loanPeriodIndex = loanStatusIndex;
  for (var period in loanPeriods) {
    dashboard.getRange(1, loanPeriodIndex).setValue(period);
    dashboard.getRange(2, loanPeriodIndex).setValue(loanPeriods[period]);
    loanPeriodIndex++;
  }

  // Outstanding Amount (assuming you have a date column to track when each loan was made)
  
  dashboard.getRange("Z1").setValue("Outstanding Amount"); // Assuming Z column for this example
  dashboard.getRange("Z2").setValue(outstandingAmount);

  // Average Repayment Time
  var repaymentTimeIndex = loanPeriodIndex;
  for (var time in repaymentTimes) {
    dashboard.getRange(1, repaymentTimeIndex).setValue(time);
    dashboard.getRange(2, repaymentTimeIndex).setValue(repaymentTimes[time]);
    repaymentTimeIndex++;
  }

  dashboard.getRange("D1").setValue("Comments for Decline");
  dashboard.getRange("D2").setValue(comments.join(", "));
}


function setupDashboardUpdateTrigger() {
  // Delete existing triggers (to avoid duplicates)
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "updateDashboard") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create a new daily trigger for updating the dashboard
  ScriptApp.newTrigger('updateDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(0) // Midnight
    .create();
}

function logChanges(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var user = Session.getActiveUser().getEmail();
  var timestamp = new Date();

  var logsSheet = e.source.getSheetByName("Logs");
  if (!logsSheet) {
    logsSheet = e.source.insertSheet("Logs");
  }

  logsSheet.appendRow([timestamp, user, 'Changed cell ' + range.getA1Notation() + ' from "' + e.oldValue + '" to "' + e.value + '"']);
}
function setupLoggingTrigger() {
  ScriptApp.newTrigger('logChanges')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}


