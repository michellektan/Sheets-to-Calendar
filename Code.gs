function update(){ // grays out rows if the event has passed
  var sheet = SpreadsheetApp.getActiveSheet();
  if(sheet.getName() == 'Main'){
  var startRow = 3;
  var startCol = 1;
  var dataRange = sheet.getRange(startRow, startCol, sheet.getLastRow()-startRow+1 , sheet.getLastColumn()); 
  var data = dataRange.getValues();
  for(var k = 0; k < data.length; ++k){
    var row = data[k];
        if(row[4] == 'TBD' || row[4] == 'TBA'){ //ignore if event hasn't been created
    }
    else{
    var eventDate = Utilities.formatDate(new Date(row[4]), "America/Los_Angeles", "yyyy-MM-dd HH:mm:ss");
    var currentDate = Utilities.formatDate(new Date(), "America/Los_Angeles", "yyyy-MM-dd HH:mm:ss");
    if(eventDate.valueOf() < currentDate.valueOf()){ //compares each event's end date with current date
      var range = sheet.getRange(startRow+k, startCol,1, sheet.getLastColumn());
      range.setBackground("#A9A9A9"); //sets gray if event has passed
  }
}
}
}
}

function sheetsToCal(){ //pulls events from spreadsheet to calendar
  var sheet = SpreadsheetApp.getActiveSheet();
  if(sheet.getName() == 'Main'){
  var calendar = CalendarApp.getCalendarById('cal'); //replace 'cal' with the calendar ID to your synced Google Calendar
  var lastRow = sheet.getLastRow();
  var startRow = 3;
  var startCol = 1;
  var numRows = sheet.getLastRow();
  var numColumns = sheet.getLastColumn();
  
  var dataRange = sheet.getRange(startRow, startCol, numRows-startRow+1, numColumns); 
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) { //sets the event details
    try{
    var row = data[i];
    var name = row[1]; 
    var startDate = new Date(row[3]);
    var endDate = new Date(row[4]);
    var location = row[5];
    var publicEvent = row[7]; //put on calendar if public
    var notes = row[8]; //event description
    var eventId = row[10]; //ID generated is placed in the last col
    var updatedEvent;
    }
    catch(e){
      sheet.getRange(startRow+i, 12).setValue("error 1");
    }
    if(publicEvent == "Yes"){
      if(eventId == 0 || eventId == null){ //if the event hasn't been created yet, create it
        try{
      var event = calendar.createEvent(name, startDate, endDate, {location: location, description: notes});
      var range = sheet.getRange(startRow+i, 11);
      range.setValue(event.getId()); //save the event's ID
        }
        catch(e){
          sheet.getRange(startRow+i, 12).setValue("error 2");
        }
      }
      else{ //if the event has been created, only update it
        try{
          updatedEvent = calendar.getEventById(eventId);
          updatedEvent.setDescription(notes);
          updatedEvent.setTitle(name);
          updatedEvent.setLocation(location);
          updatedEvent.setTime(startDate, endDate);
          sheet.getRange(startRow+i, 12).clearContent();
      }
        catch(e){
          sheet.getRange(startRow+i, 12).setValue("error 3").setFontColor("red");
      }
      }
    }
   else{
     if(eventId != 0 && eventId != null){ //removes event from sheet and calendar if it's not public
       try{
       updatedEvent = calendar.getEventById(eventId);
       updatedEvent.deleteEvent();      
       sheet.getRange(startRow+i, 11).clearContent(); //delete eventID
     }
       catch(e){
         sheet.getRange(startRow+i, 12).setValue("error 4");
   }
  }
 }
}
  }
}
