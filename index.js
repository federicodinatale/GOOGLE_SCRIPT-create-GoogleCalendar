/**
 * please note that I was running this script where I used to work. 
 * They sent me the schedule on a Single Page Application with poor frontend taste. 
 * 
 * Therefore I decided to creare a scritp that automatically will creare a nicer schedule on my Google Calendaraccount
 *
 * When I pasted the schedule from the SPA, the data was not aligned. Thefore I created the functions alignTable() and format() in order to align *as follow: 
 *
 * COLUMN A   |  COLUMN B    | COLUMN C   |  COLUMN D
 * date       |   time start | time end   |   name
 *
 * 
 * 
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Create New Shift')
    .addItem('Align Table', 'alignTable')
    .addItem('New Shift', 'createShift')
    .addToUi(); 
}
 
//in case we need to ask email to user. We also need to add on ui
function showEmail() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt('Insert Your Email');
  let email = response.getResponseText();
  return createShift(email);
}
 
 
function doGet() {
  return HtmlService.createHtmlOutput();
}
 
function alignTable() {
 
  for (let i =1; i <50; i ++) {
 
  valueA =  SpreadsheetApp.getActiveSpreadsheet().getRange(`a${i}`).getValue();
  valueB =  SpreadsheetApp.getActiveSpreadsheet().getRange(`b${i}`).getValue();
  valueC =  SpreadsheetApp.getActiveSpreadsheet().getRange(`c${i}`).getValue();
 
  if (valueC == "") {
 
    valueC =  SpreadsheetApp.getActiveSpreadsheet().getRange(`c${i}`).setValue(valueB);
    valueB =  SpreadsheetApp.getActiveSpreadsheet().getRange(`b${i}`).setValue(valueA);
    valueA =  SpreadsheetApp.getActiveSpreadsheet().getRange(`a${i}`).setValue("");
 
    console.log( )
 
  }
  }
 
  format();
  changeString();
  changeTime();
}
 
 
function format() {
 
  let check = true;
  let count = 1;
 
  while(check) {
     
    valueAcurrent = SpreadsheetApp.getActiveSpreadsheet().getRange(`a${count}`).getValue();
    valueANext = SpreadsheetApp.getActiveSpreadsheet().getRange(`a${count+1}` ).getValue();
 
    console.log("valueAcurrent: "+  valueAcurrent);
    console.log("valueANext: "+  valueANext);
 
    valueBcurrent = SpreadsheetApp.getActiveSpreadsheet().getRange(`b${count}`).getValue();
    console.log("valueB: " + valueBcurrent)
 
    if(valueBcurrent != "") {
 
      if (valueANext == "") {
 
        valueANext = SpreadsheetApp.getActiveSpreadsheet().getRange(`a${count+1}`).setValue(valueAcurrent)
        count ++;
      } else {
        console.log("piena")
        count++;
      }
  
 
    } else {
        check = false;
    }
 
  }
}
 
function changeString() {
  let check = true;
  let count = 1;
 
  while(check) {
 
    valueAcurrent = SpreadsheetApp.getActiveSpreadsheet().getRange(`a${count}`).getValue();
 
    if (valueAcurrent != "") {
      let newDate =  new Date(valueAcurrent);
      console.log(typeof Utilities.formatDate(valueAcurrent,"GMT", 'dd-MM'))
      count ++;
    } else {
      check = false;
    }
  }
}
 
function changeTime() {
  SpreadsheetApp.getActiveSpreadsheet().insertColumnAfter(2);
 
  let check = true;
  let count = 1;
 
  while (check) {
    valueBcurrent = SpreadsheetApp.getActiveSpreadsheet().getRange(`b${count}`).getValue();
    if (valueBcurrent != "") {
      
      valueTime = valueBcurrent.split("-")
      console.log(valueTime)
 
      valueBcurrent = SpreadsheetApp.getActiveSpreadsheet().getRange(`b${count}`).setValue(valueTime[0])
      valueCcurrent = SpreadsheetApp.getActiveSpreadsheet().getRange(`C${count}`).setValue(valueTime[1])
 
      count ++;
 
    } else {
      check = false;
    }
 
  }
 
}
 
function createShift()  {
 
  let months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  //let shift = ["Phone", "Picklist", "Break", "Meeting", "1to1"]
  
  for ( let i = 1; i <= 100; i ++) {
  
  
  //full date
  let date = SpreadsheetApp.getActiveSheet().getRange(`a${i}`).getValue();
  
  if (date != "") {
  
    console.log("dateDay: " + date)
  
    //year 
    let dateYear = date.getFullYear();
    console.log("dateYear: " + dateYear);
    
    //month
    let dateMonth = date.getMonth()
    let month = months[dateMonth];
    console.log("dateMonth: " + month)
    
    //day
    let dateDay = date.getDate()
    console.log("dateDay: " + dateDay);
    
    
    
    //HOURS - MINUTES
    
    //start
    let startTime = SpreadsheetApp.getActiveSheet().getRange(`b${i}`).getValue();
    
    //hours
    let startHour = startTime.getHours()
    console.log("start Hours: " + startHour)
    
    //minutes
    let startMinutes = startTime.getMinutes()
    console.log("start Minutes: " + startMinutes)
    
    
    //end
    let endTime = SpreadsheetApp.getActiveSheet().getRange(`c${i}`).getValue();
    
    //hours
    let endHour = endTime.getHours()
    console.log("end Hours: " + endHour)
    
    //minutes
    let endMinutes = endTime.getMinutes()
    console.log("end Minutes: " + endMinutes)
    
    
    let shift = SpreadsheetApp.getActiveSheet().getRange(`d${i}`).getValue();
    console.log(shift)
    
    let tstart = new Date(`${month} ${dateDay}, ${dateYear} ${startHour}:${startMinutes}:00 GMT+0200`);
    let tstop = new Date(`${month} ${dateDay}, ${dateYear} ${endHour}:${endMinutes}:00 GMT+0200`);
    
    console.log(tstart);
    console.log(tstop);
 
      let ss = SpreadsheetApp.getActiveSpreadsheet();
      
    let email = ss.getSheetByName('Dashboard').getRange("e5").getValues()[0][0]
    console.log(email);
    
    let event = CalendarApp.getDefaultCalendar().createEvent(shift, tstart, tstop);
    console.log('Event ID: ' + event.getId());
    let calendar = CalendarApp.getCalendarById(email);
 
    console.log("time script:" + Session.getScriptTimeZone());
    console.log("Calendar Script:" + CalendarApp.getTimeZone());
 
 
     let calEvents = CalendarApp.getCalendarById(email).getEvents(tstart,tstop);
      console.log(calEvents.length);
    
      for (let j = 0; j < calEvents.length; j++) {
      let orari = calEvents[j];
      let title = calEvents[j].getTitle();
 
      if(title == "Email") {
        orari.setColor("10")
      } else if (title == "Ready" || title == "Trabaja") {
        orari.setColor("2")
        orari.setTitle("Ready")
 
      } else if (title == "Break" || title == "Descanso") {
        orari.setColor("8");
        orari.setTitle("Break")
      } else if (title == "Meeting" || title == "1to1") {
        orari.setColor("9")
      } else {
        orari.setColor("11")
      }
      }
 
    } else {
    console.log("finito")
  }
  }   
} 
 
 //just to debug and check if the time script is the same as the Calendar script
function getTime() {
      console.log("time script:" + Session.getScriptTimeZone());
    console.log("Calendar Script:" + CalendarApp.getTimeZone());
}
 