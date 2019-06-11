// go from texercise strings to what we need to set the time
function getTime(string){
  if(string == "Noon"){return [12,0,0];}
  var arr = string.split(':');

  if(arr.length == 1){
    if(arr[0].toString(10).indexOf('p') > -1){return [parseInt(arr[0], 10)+12,0,0];}
    if(arr[0].toString(10).indexOf('a') > -1){return [parseInt(arr[0], 10),0,0];}
  }
  if(arr.length == 2){
    if(arr[0].indexOf('12') > -1){return [parseInt(arr[0], 10),parseInt(arr[1], 10),0];}
    if(arr[1].indexOf('p') > -1){return [parseInt(arr[0], 10)+12,parseInt(arr[1], 10),0];}
    if(arr[1].indexOf('a') > -1){return [parseInt(arr[0], 10),parseInt(arr[1], 10),0];}
  }
  Logger.log('error: ' + string);
}



function createEventSeries(first_date, last_date, day, data) {
  var times = data[0].split(' ');
  
  // make variable for the start date/time
  var start_date = new Date(first_date);
  var [start_h, start_m, start_s] = getTime(times[0]);
  start_date.setHours(start_h, start_m, start_s);
  
  // make variable for the end date/time
  var end_date = new Date(first_date);
  var [end_h, end_m, end_s] = getTime(times[2]);
  end_date.setHours(end_h, end_m, end_s );
  
  var recur = [CalendarApp.Weekday.SUNDAY, CalendarApp.Weekday.MONDAY, CalendarApp.Weekday.TUESDAY, CalendarApp.Weekday.WEDNESDAY, CalendarApp.Weekday.THURSDAY, CalendarApp.Weekday.FRIDAY, CalendarApp.Weekday.SATURDAY]
 
  // create event series 
  var eventSeries = CalendarApp.getCalendarById('tsmt2uulcn0irejcsgkup8tl2k@group.calendar.google.com').createEventSeries(data[1] + " " + data[3],
    start_date,
    end_date,
    CalendarApp.newRecurrence().addWeeklyRule()
        .onlyOnWeekdays([recur[day]])
        .until(last_date),
    {location: data[2]});
  
  // update color based on location
  if (data[2].indexOf('RSC') > -1){ var color = 9}
  else { var color = 7}
  eventSeries.setColor(color)
  
  // log this!
  Logger.log('Event Series ID: ' + eventSeries.getId());
  
}  
  


function onOpen() {
  // Add a custom menu to the spreadsheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  var numCols = sheet.getLastColumn();
  
  var dataRange = sheet.getDataRange().getValues();
  var data = sheet.getRange(1, 1, dataRange.length, 5).getValues();
  
  var days = new Array("Sundays", "Mondays", "Tuesdays", "Wednesdays", "Thursdays", "Fridays", "Saturdays");

  
  for(i in data){
    //get dates
    if(i == 3){
      var first = new Date(data[i][0]);
      var last = new Date(data[i][1]);
      Logger.log('first: ' +  first);
      Logger.log('last: ' +  last);
    }
    
    //read in the rest of the data
    if(i >= 5){
      // if it begins with a letter, get new week day
      if(/^[a-z].*/i.test(data[i][0])){
        var day = days.indexOf(data[i][0]);
      }
      // if it begins with a number, this is a new event
      if(/^[0-9].*/i.test(data[i][0])){
        createEventSeries(first, last, day, data[i])
      }
    }
  }
  
}

  
