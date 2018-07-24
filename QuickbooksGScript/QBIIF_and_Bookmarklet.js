/*
  This script was created from a Google Sheet where we input payroll information in a format that works
  for us. It's not standard or anything, but it captures what we need. This script generates:
  
  A file for QuickBooks - this file needs to be downloaded as "plain text", then I run a simple perl script 
  to clean it up so it can then be imported into QuickBooks
  
  A Bookmarklet for First Citizens Bank business ACH - I copy the text of this file to a bookmarket that I then run
  while on the bank's website, and it fills in all of the values appropriately.
  
  An email to each person with their pay

*/


// I'm using "XXXXXX" in place of tabs because the tabs don't make it to the download. After downloading, I just have
// to replace "XXXXXX" with a tab character and I'm all good.

// 1/11/15 - adding capability that creates a javascript bookmarklet each time this is run. That bookmarklet is to be run
// in Chrome on the BOB Advantage First Citizen's website after loading the Gulf Breeze Software Partners ACH template. It will
// fill in the correct amount for each person.



function createIif() {
  // debug is "yes" for debug and "no" for non-debug. This is so I have to think about it a bit before just 
  // clicking.
  var debug = Browser.msgBox('Greetings', 'For Debug, press YES. To run the script for real, press NO.', Browser.Buttons.YES_NO);
  //var debug = "yes";
  var newDoc = DocumentApp.create("_QB_IIF_Export_" + Utilities.formatDate(new Date(), "GMT-05:00", "MMddYYYY"));
  DriveApp.getFileById(newDoc.getId()).setStarred(true);
  var body = newDoc.getBody();
  
  var bmDoc = DocumentApp.create("_PAYROLL_Bookmarklet_" + Utilities.formatDate(new Date(), "GMT-05:00", "MMddYYYY"));
  DriveApp.getFileById(bmDoc.getId()).setStarred(true);
  var bmBody = bmDoc.getBody();
  
  var bmBodyText = 'javascript:(function(){var ifr=document.getElementById("ifrm");';
  
  // This is how to set the body text. I don't BELIEVE it'll be that hard. I may be wrong, but we'll do it.
  //body.setText(text);
  
  var re = /(dave|^frank$|^exactly$|^Consultant$|^\s+$)/;
  
  var bodyText = "!TRNSXXXXXXTRNSIDXXXXXXTRNSTYPEXXXXXXDATEXXXXXXACCNTXXXXXXNAMEXXXXXXAMOUNTXXXXXXDOCNUMXXXXXXMEMO\n!SPLXXXXXXSPLIDXXXXXXTRNSTYPEXXXXXXDATEXXXXXXACCNTXXXXXXAMOUNTXXXXXXMEMO\n!ENDTRNS\n";
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  // this is the date that the bookmarklet will set the First Citizens ACH to run.
  var runDate = "";
  
  var i = 0;
  for (var j = 0; j <= numRows - 1; j++) {
    var iifLineIndex = 0;
    var line = values[j];
    var iifLines = new Array();
    
    
    // If amount to be paid is $0 or it's Dave, move on to the next line in the spreadsheet, but still create the bookmarklet
    // entry, UNLESS it's a check to Dave. So we need to do another test.
    if (((line[0] == "") || line[11] <= 0) || line[1].equals("")) {
      continue;
    }
    if (re.test(line[1])) {
      var davecheckre = /check/;
      if (davecheckre.test(line[2])) {
        // if the description contains the word "check", it *should* mean that we're sending an actual check (normally to Dave), so 
        // don't include it in the bookmarklet, since it won't be in the ACH.
        continue;
      } else {
        bmBodyText = bmBodyText + createBookmarklet(line[1],line[11]);
        continue;
      }
    }
      
    bmBodyText = bmBodyText + createBookmarklet(line[1],line[11]);
    
    if (runDate == "") {
      // all of the entries that are processed have the same date, so just pick the first one.
      runDate = Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy");
    }
    //line[0].setFullYear(line[0].getFullYear()-1);
    
    // Use the following when paying from Huntington account 
    //iifLines[iifLineIndex++] = 'TRNSXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXHuntington CheckingXXXXXX' + line[2] + 'XXXXXX-' + line[11] + 'XXXXXXACHXXXXXX' + line[1];
    
    
    // Use this one when paying from First Citizens
    iifLines[iifLineIndex++] = 'TRNSXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXFirst CitizensXXXXXX' + line[2] + 'XXXXXX-' + line[11] + 'XXXXXXACHXXXXXX' + line[1];
    
    // create to:, subject and body of email
    var emailTo = new String(line[1]).replace(/ /g,".") + "@mydomain.com";
    // to fix IV's personal entry that we only make once per year.
    var mailSubject = "Check for " + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy");
    var mailBody = Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + " check will be for $" + line[11] + "\n\nThe line items are:\n\n";
    
    if (line[3] > 0) {
      iifLines[iifLineIndex++] = 'SPLXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXOutside ServicesXXXXXX' + line[3] + 'XXXXXX' + line[1] + ' salary/minimums';
      iifLines[0] = iifLines[0] + ' minimums';
      mailBody = mailBody + "Minimums = $" + line[3] + "\n";
      
    } else {
      // I believe this is redundant
      iifLines[1] = null;
    }
    
   
    if ((line[4] != "") && (line[4] != " ")) {
      // hourly
      iifLines[iifLineIndex++] = 'SPLXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXOutside ServicesXXXXXX' + line[4] + 'XXXXXX' + line[1] + ' hourly';
      iifLines[0] = iifLines[0] + ' +hourly';
      mailBody = mailBody + "Hourly pay = $" + line[4] + "\n\n        Hourly Line Items:\n\n";
      
      
      var lineItems = false;
          
      for (var k = j; k < numRows; k++) {
        var hrlyLine = values[k];
        if ((hrlyLine[1].toLowerCase() == line[1].toLowerCase()) && (hrlyLine[0] == "") && (hrlyLine[3] > 0) && ((hrlyLine[2].indexOf("hours") > 0) || (hrlyLine[2].indexOf("hourly") >= 0))) {
          // We've found one of the hourly line items for this person (line[1] has his/her name)
              
          
          lineItems = true;
              
          mailBody = mailBody + "        " + hrlyLine[2] + ":   $" + hrlyLine[3] + "\n\n";
        }
      }
    } else {
      iifLines[4] = null;
    }
    if ((line[5] != "") && (line[5] != " ")) {
      // q1 bonus
      iifLines[iifLineIndex++] = 'SPLXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXOutside ServicesXXXXXX' + line[5] + 'XXXXXX' + line[1] + ' +q1 bonus pmt';
      iifLines[0] = iifLines[0] + ' +q1 bonus pmt';
      mailBody = mailBody + "q1 bonus payment = $" + line[5] + "\n";
    } else {
      iifLines[5] = null;
    }
    if ((line[6] != "") && (line[6] != " ")) {
      // q2 bonus
      iifLines[iifLineIndex++] = 'SPLXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXOutside ServicesXXXXXX' + line[6] + 'XXXXXX' + line[1] + ' +q2 bonus pmt';
      iifLines[0] = iifLines[0] + ' +q2 bonus pmt';
      mailBody = mailBody + "q2 bonus payment = $" + line[6] + "\n";
    } else {
      iifLines[6] = null;
    }
    if ((line[7] != "") && (line[7] != " ")) {
      // q3 bonus
      iifLines[iifLineIndex++] = 'SPLXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXOutside ServicesXXXXXX' + line[7] + 'XXXXXX' + line[1] + ' +q3 bonus pmt';
      iifLines[0] = iifLines[0] + ' +q3 bonus pmt';
      mailBody = mailBody + "q3 bonus payment = $" + line[7] + "\n";
    } else {
      iifLines[7] = null;
    }
    if ((line[8] != "") && (line[8] != " ")) {
      // q4 bonus
      iifLines[iifLineIndex++] = 'SPLXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXOutside ServicesXXXXXX' + line[8] + 'XXXXXX' + line[1] + ' +q4 bonus pmt';
      iifLines[0] = iifLines[0] + ' +q4 bonus pmt';
      mailBody = mailBody + "q4 bonus payment = $" + line[8] + "\n";
    } else {
      iifLines[8] = null;
    }
    if ((line[9] != "") && (line[9] != " ")) {
      
      mailBody = mailBody + "Expenses total = $" + line[9] + "\n\n        Expense Line items:\n\n";
      
      // expenses - look further down in the spreadsheet for the expense line items
      var lineItems = false;
      for (var k = j; k < numRows; k++) {
        var expLine = values[k];

        if ((expLine[1].toLowerCase() == line[1].toLowerCase()) && (expLine[0] == "") && (expLine[3] > 0) && (expLine[2].indexOf("exp") >= 0)) {
          // We've found one of the expense line items for this person (line[1] has his/her name)
          //Commented out on 1/27
          //expLine[3] = -1 * expLine[3];
          iifLines[iifLineIndex++] = 'SPLXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXTravel ReimbursementsXXXXXX' + expLine[3] + 'XXXXXX' + line[1] + ' ' + expLine[2];
          lineItems = true;
          
          mailBody = mailBody + "        " + expLine[2] + ":   $" + expLine[3] + "\n";
        }
      }
      if (!lineItems) {
        iifLines[iifLineIndex++] = 'SPLXXXXXX' + i++ + 'XXXXXXCHECKXXXXXX' + Utilities.formatDate(line[0], "GMT", "MM/dd/yyyy") + 'XXXXXXTravel ReimbursementsXXXXXX' + line[9] + 'XXXXXX' + line[1] + ' expenses';
      }
      
      iifLines[0] = iifLines[0] + ' +expenses';
    } else {
      iifLines[9] = null;
    }
    iifLines[iifLineIndex++] = 'ENDTRNS';
    
    // send the email
    
    sendEmail(emailTo,mailSubject,mailBody,debug);
        
    for (var k = 0; k < iifLines.length; k++) {
      if (iifLines[k] != null) {
        bodyText = bodyText + '\n' + iifLines[k];
      }
    }
  }
  body.setText(bodyText);
  
  bmBody.setText(bmBodyText);
}

/**
 * This is the function that builds the JavaScript bookmarklet
**/
function createBookmarklet(theperson,theamount) {
  // the resulting line needs to read:
  //
  // var amount1 = document.getElementById("ifrm").contentDocument.getElementById("amount1"); amount1.value = "theamount"; amount1.onchange();
  //
  // The entire bookmarklet will look like:
  //
  // javascript: (function() { var amount1 = document.getElementById("ifrm").contentDocument.getElementById("amount1"); amount1.value = "theamount"; amount1.onchange();} )();
  
  var position = 0;
  var thisLine = "";
  var templateLine = 'var amountXXX=ifr.contentDocument.getElementById("amountXXX");amountXXX.value="YYY";amountXXX.onchange();';
  switch(theperson.toLowerCase()) {
    
    case "first person" : position = "1";break;
      
    case "second person" : position = "2";break;
      
    case "third person" : position = "3";break;
      
    case "fourth person" : position = "4";break;
    
    case "fifth person" : position = "5";break;
    
    case "sixth person" : position = "6";break;  
   
    case "seventh person" : position = "7";break;
          
    case "eighth person" : position = "8";break;
      
    case "ninth person" : position = "9";break;
      
    
  }
  
  if (position > 0) {
    thisLine = templateLine.replace(/XXX/g,position);
    thisLine = thisLine.replace(/YYY/g,theamount);
  }
  return(thisLine);
  
  
};


/**
 * Adds a custom menu to the active spreadsheet
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "MakeIIF",
    functionName : "createIif"
  }];
  sheet.addMenu("GBSP", entries);
};


// This is called by createIif() above to send emails to folks.
function sendEmail(email,subject,body,debug) {
  
  if (debug == "yes") {
    
    Logger.log(email);
    Logger.log(subject);
    Logger.log(body);
  } else if (debug == "no") {
    MailApp.sendEmail(email, subject, body, {cc: "myname@mydomain.com,someoneelse@mydomain.com"});
  }
  
}
