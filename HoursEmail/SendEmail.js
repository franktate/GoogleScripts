/*
  This script was created from within a Google Sheet with a form. Every Friday, a link to the form is sent 
  to all consultants to have them enter their hours.
  
*/

// send an email with a better subject than the default when someone submits a form response
function onFormSubmit(e) {
  // e.values contains the values, in the order submitted.
  // e.namedValues contains a hash of the name,value pairs.
  // anthony wants the format to be "Hours - Date - Client - SOW"
  
  // This quit working on 11/23/2016. The fix was to use e.values[1] rather than e.namedValues["Username"].
  // gotta love impromptu changes in the API!
  //MailApp.sendEmail(e.namedValues["Username"], "w/e: " + e.values[2] + " - " + e.values[4] + " hours - " + e.values[3] + " - " + e.values[12], 
  //                  "Notes: " + e.values[8]);
  
  MailApp.sendEmail(e.values[1], "w/e: " + e.values[2] + " - " + e.values[4] + " hours - " + e.values[3] + " - " + e.values[12], 
                    "Notes: " + e.values[8]);
    
  MailApp.sendEmail("frank.tate@gulfsoft.com,kelly.pryor@gulfsoft.com", e.values[4] + "hrs," + e.values[1] + "," +e.values[12] + "," + e.values[2] + ",Cust:" + e.values[3], 
                    "SOW: " + e.values[12] + "\nweek ending: " + e.values[2] + "\nhours: " + e.values[4] + "\nCustomer: " + e.values[3] + 
                    "\nConsultant: " + e.values[1] + "\nExpenses: " + e.values[5] + "\nTask Name: " + e.values[7] + "\nNotes: " + e.values[8]);
  
  sendJSONhours(e);
  
}

// This function authenticates to my home server to send data each time someone logs their hours.
// The authentication works perfectly; You can get the encoded string just using "inspect element" in Firefox when 
// testing locally. 
function sendJSONhours(e) {
  var headers = {
    "Authorization" : "Basic Base64EncodedUserIDAndPassword"
  };
  
  // payload gets turned directly into $_POST in PHP. So I just need to build the array of data that I care about.
  
  var payload =
   {
     "timestamp": e.values[0],
     "username" : e.values[1].toString(),
     "weekending" : e.values[2],
     "customername" : e.values[3],
     "hoursworked" : e.values[4],
     "expenses" : e.values[5],
     "ponum" : e.values[6],
     "taskname" : e.values[7],
     "notes" : e.values[8],
     "sowname" : e.values[12],
     "onsiteorremote" : e.values[14],
     "nextweek" : e.values[15]
   };

  var params = {
    "method":"POST",
    "headers":headers,
    "payload":payload
  };

  // Send the data to a remote application that stores that data. The one I send it to just stores the data
  // in a MariaDB table to be processed by another application.
  var reponse = UrlFetchApp.fetch("http://yourhost/yourpath/yourfile.php", params);
  
}

// a simple function to re-send any hours that were somehow missed when the form was submitted.
// to use it, first select the row that needs to be sent, then select the GBSP menu option and
// select "SendOneEntry"
function failedFormSubmit() {
  
  var theRange = SpreadsheetApp.getActiveRange();
  //Logger.log(SpreadsheetApp.getActiveRange().getNumColumns());
  var theData = new Object;
  theData.values = theRange.getValues()[0];
  //Logger.log(theData);
  sendJSONhours(theData);
  MailApp.sendEmail("frank.tate@gulfsoft.com,kelly.pryor@gulfsoft.com", "RETRYING MANUALLY: " + theData.values[4] + "hrs," + theData.values[1] + "," +theData.values[12] + "," + theData.values[2] + ",Cust:" + theData.values[3], 
                    "SOW: " + theData.values[12] + "\nweek ending: " + theData.values[2] + "\nhours: " + theData.values[4] + "\nCustomer: " + theData.values[3] + 
                    "\nConsultant: " + theData.values[1] + "\nExpenses: " + theData.values[5] + "\nTask Name: " + theData.values[7] + "\nNotes: " + theData.values[8]);
  
}


// Adds a GBSP menu item with one option: "SendOneEntry".
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "SendOneEntry",
    functionName : "failedFormSubmit"
  }];
  sheet.addMenu("GBSP", entries);
};