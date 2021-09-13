
function onSubmit(e) {
  
  var errorFlag = 0;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log('Could not obtain lock after 10 seconds.');
    console.log('Could not obtain lock after 10 seconds.');
  }
  
  //var formDestId = FormApp.getActiveForm().getDestinationId();
  var ss = SpreadsheetApp.openById("1_Q2lAqpuOz7pK1MZ23DrwuJ3tnT-29OgE6DT6ypPHBQ");
  
  SpreadsheetApp.setActiveSpreadsheet(ss);
  shFormResponses = ss.getSheetByName('Form Responses 1');
  
  var responses = FormApp.getActiveForm().getResponses();
  
  var length = responses.length;
  
  var lastResponse = responses[length-1];
  var formValues = lastResponse.getItemResponses();
  var length = lastResponse.getItemResponses().length;
  var formFindValue = formValues[0].getResponse();
  var priority = 0;
  
  Logger.log("length = " + length + ". lastResponse = " + lastResponse + ". formValues = " + formValues + ". formFindValue = " + formFindValue);
  
  
  
  Logger.log("Form find value: " + formFindValue);
  //Logger.log("Form destination id : " + FormApp.getActiveForm().getDestinationId());
  
  var currentRow, tktSubmitterEmail;
  
  if(formFindValue != "Case#") {
    var values = shFormResponses.getDataRange().getValues();
    //Logger.log("Form find for total values: " + values);
    for(var j=0, jLen=values.length; j<jLen; j++) {
      for(var k=0, kLen=values[0].length; k<kLen; k++) {
        var find = values[j][k];
        //Logger.log(find);
        if(find == formFindValue) {
          //Logger.log([find, "row "+(j+1)+"; "+"col "+(k+1)]);
          currentRow = j+1;
          break;           
        }
      }
    }
  }
  
  ss.setActiveSheet(shFormResponses);
  
  var ticketNumber;
  var ticketCounter;
  var row;
  var etrolControlsServiceEmail1, etrolControlsServiceEmail2, etrolControlsServiceEmail3 = null, timestamp, ticketNumberLocation;
  var editFlag = 0;
  var numberAsNumber2;
  var caseType;
  var meetingDate;
  var meetingTime;
  var meetingFinish;
  var duplicateCaseNum;
  var caseStatus;
  var missingDuplicateFlag = 0;
  var customerName = "";
  
  var months = {
      '1': "January", 
      '2': "February", 
      '3': "March", 
      '4': "April", 
      '5': "May", 
      '6': "June", 
      '7': "July", 
      '8': "August", 
      '9': "September", 
      '10': "October", 
      '11': "November", 
      '12': "December"
  };
    
  //var responses = FormApp.getActiveForm().getResponses();
  //var responseId = responses[0].getRespondentEmail();
  
  if(currentRow) {
    
    tktSubmitterEmail = shFormResponses.getRange(currentRow, 4).getValues();
    
    Logger.log("1. Submitter email is " + tktSubmitterEmail + " currentRow " + currentRow);
  }
  
  /* Dealing with invalid entries */
  if(!currentRow && (formFindValue != "Case#")) /* We couldn't find your case so don't let them submit */
  {
    tktSubmitterEmail = e.response.getRespondentEmail();
    
    
    Logger.log("Invalid entry " + formFindValue + "getEditors returned" + tktSubmitterEmail);
    errorFlag = 2;
  }
  
  if(currentRow && (formFindValue != "Case#")) /* Email mismatch between submitter and Account manager */
  {
    tktSubmitterEmail = e.response.getRespondentEmail();
    Logger.log("Ticket submitter email is: " + tktSubmitterEmail + " and the original submitter email is: " + shFormResponses.getRange(currentRow, 1).getValues().toString());
    
    var tktSEemail = null;
    tktSEemail = shFormResponses.getRange(currentRow, 7).getValues().toString();
    
    if (tktSEemail) {   /* Checking if PM is changing tkt status */
      if (tktSubmitterEmail != 'jay.bhaskar@gmail.com' && tktSubmitterEmail != 'dkumets@fastly.com' && tktSubmitterEmail != tktSEemail && tktSubmitterEmail != shFormResponses.getRange(currentRow, 1).getValues().toString())
      //if (tktSubmitterEmail != shFormResponses.getRange(currentRow, 1).getValues().toString())
      {
        Logger.log("Not your ticket to edit " + formFindValue);
        errorFlag = 1;
      }
    } else {
      if (tktSubmitterEmail != 'jay.bhaskar@gmail.com' && tktSubmitterEmail != 'dkumets@fastly.com' && tktSubmitterEmail != shFormResponses.getRange(currentRow, 1).getValues().toString())
      //if (tktSubmitterEmail != shFormResponses.getRange(currentRow, 1).getValues().toString())
      {
        Logger.log("Not your ticket to edit " + formFindValue);
        errorFlag = 1;
      }
    }
  }
  
  if((currentRow === undefined || currentRow === null) && formFindValue == "Case#") { /* This is a new entry */
    /* Checking if you're allowed to pick the A/c manager */
    tktSubmitterEmail = e.response.getRespondentEmail();
    if (tktSubmitterEmail != 'jay.bhaskar@gmail.com' && tktSubmitterEmail != 'dkumets@fastly.com' && tktSubmitterEmail != formValues[3].getResponse()) // The index for formValues will change depending on whether
    //if (tktSubmitterEmail != 'bjayaraman@fastly.com' && tktSubmitterEmail != formValues[3].getResponse()) // The index for formValues will change depending on whether
    {
      errorFlag = 3; 
    }
  }  
  
  
  if(currentRow && (formFindValue != "Case#")) /* Extract portion after Case# and validate value -- Not implemented yet*/
  {
    var localValue = formFindValue[0];
  }
  
  /* Dealing with invalid entries ends */
  
  if(errorFlag == 0) {
    
    if(currentRow && (formFindValue != "Case#") ) { /* 0. Someone is editing an existing response so get ticket number */
      
      Logger.log("formFindValue" + formFindValue);
      ticketNumber = shFormResponses.getRange(currentRow, 2).getValues();
      row = currentRow;
      ticketCounter = currentRow;
      editFlag = 1;
      
    }
    
    var tktSheet = ss.getSheetByName('Maintenance');
    
    
    //if(currentRow && (formFindValue == "Case#")){
    if((currentRow === undefined || currentRow === null) && formFindValue == "Case#"){ /* This is a new entry */
      
      //Logger.log("Looking to get the ticket counter from sheet");
      ticketCounter = tktSheet.getRange(2, 1).getValue();
      
    }
    
    //Logger.log("Ticket counter at:" + ticketCounter); /* 2. Just finding out if ticket counter was initialized and is rolling */
    
    if (ticketCounter === undefined || ticketCounter === null || ticketCounter === NaN || ticketCounter == 0) { /* 3. This is the first time ticket counter is being initialized */
            
      ticketCounter = '1';
      
      tktSheet.getRange(2, 1).setValue(ticketCounter);
    }
    

    if(currentRow && (formFindValue == "Case#")) { /* this is the new ticket */
      
      numberAsNumber2 = Number(ticketCounter);
      
      if (row === undefined || row === null)

        row = parseInt(numberAsNumber2+1);
      
      ticketCounter = (numberAsNumber2+1).toString();
            
      ticketNumber = "Case#" + ticketCounter;
      
      
    } else if (currentRow === undefined) {
      numberAsNumber2 = Number(ticketCounter);
      ticketCounter = numberAsNumber2+1;
      row = ticketCounter;
      ticketNumber = "Case#" + ticketCounter;
    }
    
    etrolControlsServiceEmail1 = "jay.bhaskar@gmail.com",
      etrolControlsServiceEmail2 = "jay.bhaskar+duplicate@gmail.com",
        tktSubmitterEmail = e.response.getRespondentEmail(),
          //timestamp = shFormResponses.getRange(row, 1).getValues(),
          timestamp = e.response.getTimestamp(),
            ticketNumberLocation = shFormResponses.getRange(row, 3);
    
    
    ticketNumberLocation.setValue(ticketNumber);
    
    var sheetIndex = 1;
    
    if (tktSubmitterEmail == 'jay.bhaskar@gmail.com' || tktSubmitterEmail != 'dkumets@fastly.com') /* We are allowed to log tkts for someone else */
      shFormResponses.getRange(row, sheetIndex++).setValue(formValues[2].getResponse());
    else
      shFormResponses.getRange(row, sheetIndex++).setValue(tktSubmitterEmail);
    
    shFormResponses.getRange(row, sheetIndex++).setValue(ticketNumber); // By now the sheetIndex is at 3
    Logger.log("resp length =" + length);
    for (var resp = 1; resp < length; resp++, sheetIndex++) { // One long loop that could be made more modular and efficient
      //Logger.log(formValues[resp].getResponse() + " Sheet Index is at " + sheetIndex);
      if(formValues[resp].getResponse() == 'No' && formValues[resp].getItem().getTitle() == "Is this from a customer?") { /* We will follow discipline of getting y/n answers for items right before they're entered */
        //respFlag = 1;
        //Logger.log("Flag is now set");
        Logger.log("Setting value " + formValues[resp].getResponse() + " at sheetIndex " + sheetIndex + " row " + row);
        shFormResponses.getRange(row, sheetIndex).setValue(formValues[resp].getResponse());
        Logger.log("Will continue to next form entry +2 " + formValues[resp].getItem().getTitle());
        sheetIndex+=1;
        continue;
      } if(formValues[resp-1].getResponse() == 'Yes' && formValues[resp].getItem().getTitle() == "Customer Name"){
        customerName = formValues[resp].getResponse();
      }else if(formValues[resp].getResponse() == 'No') { /* We will follow discipline of getting y/n answers for items right before they're entered */
        Logger.log("Setting value " + formValues[resp].getResponse() + " at sheetIndex " + sheetIndex + " row " + row);
        shFormResponses.getRange(row, sheetIndex).setValue(formValues[resp].getResponse());
        //respFlag = 1;
        //Logger.log("Flag is now set");
        Logger.log("Will continue to next form entry +1 " + formValues[resp].getItem().getTitle());
        //sheetIndex++;
        //continue;
      }
      if(formValues[resp].getItem().getTitle() == "Case Type")
      {
        caseType = formValues[resp].getResponse();
      }
      if(formValues[resp].getItem().getTitle() == "Due date")
      {
        meetingDate = formValues[resp].getResponse();
      }
      if(formValues[resp].getItem().getTitle() == "Product Manager/Watcher Email")
      {
        etrolControlsServiceEmail3 = formValues[resp].getResponse();
      }
      if(formValues[resp].getItem().getTitle() == "[OPTIONAL] If meeting from what time?")
      {
        meetingTime = formValues[resp].getResponse();
      }
      if(formValues[resp].getItem().getTitle() == "If meeting for how long?")
      {
        meetingFinish = formValues[resp].getResponse();
      }
      if(formValues[resp].getItem().getTitle() == "Status")
      {
        caseStatus = formValues[resp].getResponse();
      }
      if(formValues[resp-1].getResponse() == 'Yes' && formValues[resp].getItem().getTitle() == 'Duplicate or related case number')
      {
        duplicateCaseNum = formValues[resp].getResponse();
        Logger.log("Duplicate case number " + duplicateCaseNum);
        var localValues = shFormResponses.getDataRange().getValues();
        var localCurrentRow = null;
        for(var jLocal=0, jLocalLen=localValues.length; jLocal<jLocalLen; jLocal++) {
          for(var kLocal=0, kLocalLen=localValues[0].length; kLocal<kLocalLen; kLocal++) {
            var find = localValues[jLocal][kLocal];
            if(find == duplicateCaseNum) {
              localCurrentRow = jLocal+1; // If ticket number not found then throw an error saying duplicate number missing
              break;           
            }
          }
        }
        if (localCurrentRow == null)
          missingDuplicateFlag = 1;
      }
      
      Logger.log("Setting value " + formValues[resp].getResponse() + " at sheetIndex " + sheetIndex + " row " + row);
      
      if (formValues[resp].getItem().getTitle() == "Priority") {
        Logger.log("Priority is " +  formValues[resp].getResponse());
        priority = formValues[resp].getResponse();
      }
      if (formValues[resp].getItem().getTitle() == 'Duplicate or related case number') { //We don't want to set case numbers not present
        Logger.log("Duplicate or related case number is " +  formValues[resp].getResponse() + " sheet index is " + sheetIndex);
        if (!missingDuplicateFlag)
          shFormResponses.getRange(row, sheetIndex).setValue(formValues[resp].getResponse());
      } //Else business as usual
      else
        shFormResponses.getRange(row, sheetIndex).setValue(formValues[resp].getResponse());
    }  
    
    
    //ticketCounter = (numberAsNumber2).toString();
    tktSheet = ss.getSheetByName('Maintenance');
    if (!editFlag)
      tktSheet.getRange(2, 1).setValue(ticketCounter);
    

    var reportedBy = e.response.getRespondentEmail();
    //var pmEmail = shFormResponses.getRange(row, 4).getValues();
    var pmEmail = etrolControlsServiceEmail3;
    Logger.log("pmEmail is " + shFormResponses.getRange(row, 4).getValues()); 
    //var priority = shFormResponses.getRange(row, 10).getValues();
    
    var subject;
    var emailBody;
    if (formFindValue == "Case#") {
      
      if(customerName != "") {
      subject =  "A product feature case for \""+ customerName + "\" has been reported on " + 
        timestamp + " " + "with ticket Number " + ticketNumber;
      } else {
        subject =  "A product feature case has been reported on " + 
        timestamp + " " + "with ticket Number " + ticketNumber;
      }
        
      
      emailBody = "To: Product Team Member " + 
        /*"\nRE: Issue reported by " + reportedBy + "." + */
        "\n\nAn issue has been reported for " + pmEmail + 
          ". Please see the details below:" + "\nTicket Number: " + 
            ticketNumber + "\nCase Type: " + 
              caseType + "\nCase Status: " + 
                caseStatus + "\nReported By: " + 
                  reportedBy + "\nPriority Level: " + priority + "\nAccount: " + 
                    customerName + "\nYou may view/edit this case using this URL - " + e.response.getEditResponseUrl()+"&entry.425022159="+"Case%23"+row;
    } else {
      
      if(customerName != "") {
      subject =  "A product feature case for \""+ customerName + "\" has been updated on " + 
        timestamp + " " + "with ticket Number " + ticketNumber;
      } else {
        subject =  "A product feature case has been updated on " + 
        timestamp + " " + "with ticket Number " + ticketNumber;
      }
      
      emailBody = "To: Product Team Member " + 
        /*"\nRE: Issue updated by " + reportedBy + "." + */
        "\n\nAn issue has been updated for " + pmEmail + 
          ". Please see the details below:" + "\nTicket Number: " + 
            ticketNumber + "\n Case Type: " + 
              caseType + "\nCase Status: " + 
                caseStatus + "\nReported By: " + 
                  reportedBy + "\nPriority Level: " + priority + "\nAccount: " + 
                    customerName + "\nYou may view/edit this case using this URL - " + e.response.getEditResponseUrl()+"&entry.425022159="+"Case%23"+row; /* We are prefilling the response */
      //Logger.log("Query string is " + e.response.getEditResponseUrl());
    }
    Logger.log("Meeting date is " + meetingDate + "Meeting time is " + meetingTime + " & duration is " + meetingFinish);
    if (meetingTime != NaN && meetingTime != null && meetingTime != "" && meetingFinish != NaN && meetingFinish != null && meetingFinish != "")
    {
      if (+String(meetingFinish).substring(0, 2) == "0" && +String(meetingFinish).substring(3, 5) == "0")
      {
        ;
      } else {
        var title = "";
        if (customerName)
            title = "Meeting for Customer " + customerName + " for " + formValues[1].getResponse();
        else
          title = "Internal meeting for feature " + formValues[1].getResponse();
        var calendar = CalendarApp.getCalendarById('fastly.com_4l3gdoaa7c6ho97qqkqe17ntqk@group.calendar.google.com');
        Logger.log("Meeting date is " + meetingDate + "Meeting time is " + meetingTime + " & duration is " + meetingFinish);
        
        /*var tempDate = new Date(String(meetingDate));
        Logger.log("Meeting date in Date obj format is " + tempDate.SetDate(1,1,2020));*/
        
        var year = +String(meetingDate).substring(0, 4);
        var month = +String(meetingDate).substring(5, 7);
        var day = +String(meetingDate).substring(8, 10);
        var mTimeHrs = +String(meetingTime).substring(0, 2);
        
        var mTimeMin = +String(meetingTime).substring(3, 5);
        var timeZone = Session.getScriptTimeZone();
        //var timeZone = Session.getTimeZone();
        
        var remainderDays = null;
        var remainderHrs = null;
        var remainderMins = null;
        var fTimeHrs = +String(meetingFinish).substring(0, 2);
        fTimeHrs = +String(Number(mTimeHrs)+Number(fTimeHrs));
        

        var fTimeMin = +String(meetingFinish).substring(3, 5);
        fTimeMin = +String(Number(mTimeMin)+Number(fTimeMin));
        var fTimeNum = Number(fTimeMin);
        var remainderHrsNum = 0;
        for(;fTimeNum>=60; fTimeNum-=60) {
          remainderHrsNum+=1
        }
       
        if(remainderHrsNum) {
          fTimeHrs = +String(Number(fTimeHrs)+remainderHrsNum);
          fTimeMin = +String(fTimeNum);
        }
        
        
        Logger.log("Year Month Day Hrs Min TimeZone " + year + "," + month + "," + day + "," + mTimeHrs + "," + mTimeMin + "," + timeZone);
        Logger.log("Finish Time Hrs Min " + fTimeHrs + "," + fTimeMin);
        
        var startTime = String(months[Number(month)])+" "+String(day)+", "+String(year) +" "+mTimeHrs+":"+mTimeMin;
        //var finishTime = String(months[Number(month)])+" "+String(day)+", "+String(year) +" "+fTimeHrs+":"+fTimeMin+" "+timeZone;
        var finishTime = String(months[Number(month)])+" "+String(day)+", "+String(year) +" "+fTimeHrs+":"+fTimeMin;
        Logger.log("Meeting start time is " + startTime);
        Logger.log("Meeting finish time is " + finishTime);
        var startDate = new Date(String(startTime));
        Logger.log("Meeting start date is " + startDate);
        var startDate = new Date(startTime);
        Logger.log("Meeting start date is " + startDate);
        
        
        if (formFindValue == "Case#") { /* New tkt */
          
          var advancedArgs = {description: title};
          
          if (etrolControlsServiceEmail3)
           advancedArgs = {description: title, location: 'here', guests:tktSubmitterEmail+','+pmEmail+','+etrolControlsServiceEmail3, sendInvites:true};
          else
            advancedArgs = {description: title, location: 'here', guests:tktSubmitterEmail+','+pmEmail, sendInvites:true};
          
          var event = calendar.createEvent(title, new Date(String(startTime)), new Date(String(finishTime)), advancedArgs);

          
          emailBody += "\n\nMeeting event created in your calendar for " + meetingDate + meetingTime + " " + timeZone;
                
          //event.addGuest(pmEmail);
          
          /*for each (eachEmail in pmEmail)
            event.addGuest(eachEmail);*/
          
          if (etrolControlsServiceEmail3 != null){
            for each (var eachEmail in etrolControlsServiceEmail3)
              event.addGuest(eachEmail);
          }
            //event.addGuest(etrolControlsServiceEmail3);
          if (customerName)
            event.setDescription("For customer: " + customerName+ " requesting feature " + caseType);
          else
            event.setDescription("Internal submission from: " + tktSubmitterEmail + " requesting feature " + caseType);
          shFormResponses.getRange(row, 15).setValue(event.getId());
        } else {
          //var events = calendar.getEventsForDay(new Date(meetingDate));
          var eventId = shFormResponses.getRange(row, 15).getValue();
          var event = calendar.getEventById(eventId);
          var advancedArgs = {description: title};
          Logger.log("Deleting event with id" + eventId + " on row " + row);
          
          if(caseStatus != 'Closed' && caseStatus != 'Duplicate/Closed' && caseStatus != 'Resolved' && caseStatus != 'No Action Needed' && caseStatus != 'Resolved into Opportunity') { // Could use better logic to find out if meeting time hasn't changed then don't do anything
            
            if (event != undefined && event != null)
              event.deleteEvent();
          }
          
          if (etrolControlsServiceEmail3)
           advancedArgs = {description: title, location: 'here', guests:tktSubmitterEmail+','+pmEmail+','+etrolControlsServiceEmail3, sendInvites:true};
          else
            advancedArgs = {description: title, location: 'here', guests:tktSubmitterEmail+','+pmEmail, sendInvites:true};
          
          event = calendar.createEvent(title, new Date(String(startTime)), new Date(String(finishTime)), advancedArgs);
          //event.addGuest(etrolControlsServiceEmail1);
          emailBody += "\n\nMeeting event created in your calendar for " + meetingDate + meetingTime + " " + timeZone;
          
          if(caseStatus != 'Closed' && caseStatus != 'Duplicate/Closed' && caseStatus != 'Resolved' && caseStatus != 'No Action Needed' && caseStatus != 'Resolved into Opportunity') { // Could use better logic to find out if meeting time hasn't changed then don't do anything
            //event.addGuest(pmEmail);
            if (etrolControlsServiceEmail3 != null){
            for each (var eachEmail in etrolControlsServiceEmail3)
              event.addGuest(eachEmail);
          }
            
            event.setDescription(customerName+caseType);
          }
          shFormResponses.getRange(row, 15).setValue(event.getId());
        }
        Logger.log("email " + etrolControlsServiceEmail1 + " customer " + customerName + "case type" + caseType);
        Logger.log("Event end time " + event.getEndTime() + "Event start time " + new Date(startTime));
      }
    }
    
    if (tktSubmitterEmail != etrolControlsServiceEmail1) // Main controller receives all emails
      MailApp.sendEmail(etrolControlsServiceEmail1, subject, emailBody);
    //MailApp.sendEmail(etrolControlsServiceEmail2, subject, emailBody);
    MailApp.sendEmail(tktSubmitterEmail, subject, emailBody);
    if (etrolControlsServiceEmail3 != null && tktSubmitterEmail != etrolControlsServiceEmail3) // Reduce spam for main controller
      MailApp.sendEmail(etrolControlsServiceEmail3, subject, emailBody);
    if (missingDuplicateFlag)
      MailApp.sendEmail(tktSubmitterEmail, "Couldn't find duplicate/related ticket", emailBody+"\n\nDuplicate/Related SE Ticket Submitted " + duplicateCaseNum + " Not found");
    MailApp.sendEmail(pmEmail, subject, emailBody);
    //Logger.log("2. Submitter email is " + tktSubmitterEmail);
    
  }
  else if (errorFlag == 1){
    
    var errEmailBody = "You can't submit this ticket no. " + formFindValue + " as you don't own it, sorry.";
    MailApp.sendEmail(tktSubmitterEmail,"Not your ticket to edit", errEmailBody);
  }
  else if(errorFlag == 2) {
    var errEmailBody = "This ticket no. wasn't found " + formFindValue + ". Please try again.";
    MailApp.sendEmail(tktSubmitterEmail,"Ticket not found", errEmailBody);
  }
  else if(errorFlag == 3) {
    var errEmailBody = "You can't pick the AM. Please try again.";
    MailApp.sendEmail(tktSubmitterEmail,"Not your ticket to edit", errEmailBody);
  }
  
  lock.releaseLock();
}

function onOpen(e)
{
}
