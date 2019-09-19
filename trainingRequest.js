// This function is what defines what approval is needed.
function returnApprovalNeeded(hours, cost, gm) {
  
  // Change the below two values to change what needs a group manager's approval
  var gmOnlyHours = 2;
  var gmOnlyCost = 100;
  
  // Change the below two values to change what needs an approval from both the group manager and the office manager.
  var bothHours = 8;
  var bothCost = 1000;
  
  if (gm === 'N/A'){
    return 'OM Only';
  } else if (hours >= bothHours || cost >= bothCost) {
    return 'OM & GM';
  } else if (hours >= gmOnlyHours || cost >= gmOnlyCost)  {
    return 'GM Only';
  } else {
    return 'Neither'; 
  }
}

// This function sets the level of approval needed, in the 'Approval Needed' column
function setApprovalNeeded(hours, cost, gm) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var approvalValue = returnApprovalNeeded(hours, cost);
  Logger.log(approvalValue);
  var col = getColByName('Approval Needed');
  Logger.log(col);
  sheet.getRange(sheet.getLastRow(), col).setValue(approvalValue);
}

// This function retrieves the data from the form submission and gets the rest of it rollin'
function onFormSubmit(e) {
  var values = e.namedValues;
  var cost = values['Total cost of Training Session/Seminar/Course Attendance'];
  var hours = values['Estimated Hours that will be charged to Office Overhead while you are attending/traveling to this event'];
  var gm = values['Who is your Group Manager?'];
  setApprovalNeeded(hours, cost, gm);
}

// This function allows another function to access the 'Approval Needed' column, even if it gets moved
function getColByName(colName) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  if (col != -1) {
    return col + 1;
  }
}
