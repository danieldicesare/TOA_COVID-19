// Function name: onEdit
//        Author: D.DiCesare
//   Description: Event that fires off any time something happens on the sheets. Conditions placed to specficially look at the Census tab and the 
//                status column. 
//                *** Please note, this is a workable soltuion that ties into a report that displays on the internet (See Wally for report name and location). 
//                While this works, it is fragile given the nature of spreadsheets. It is advised that if the spreadsheet is to be used long term, an alternative
//                solution should be explored
//    Parameters: e: spreadsheet event
//       Returns: Int: NA
//     Revisions: DCD 10/07/2020: Initial
//                DCD 10/16/2020: Added Friday 3PM logic
//                DCD 10/21/2020: Adjusted requirement to show all schools not just schools with COVID. Added Total Count column and populted
//                                Added rows 2 through 14 with base values to trick report into totalling correctly
//                                Added new tab Schools with name of school and sort order
//                DCD 11/09/2020: Added additional logging.
//                DCD 11/10/2020: Added logic for total count and e.prevVal
//                DCD 11/11/2020: New requirement to support coulmns shifting. Added logic to find the necessary columns for the report. 
function onEdit(e) {

  // Load column and range information
  var range = e.range; 
  //var currentRow = range.getRow();
  var currentCol = range.getColumn();
  var ss = SpreadsheetApp.getActiveSheet();
  var userProp = PropertiesService.getUserProperties();
  
  var rowStart = e.range.rowStart;
  var rowEnd = e.range.rowEnd;
  var columnStart = e.range.columnStart;
  var columnEnd = e.range.columnEnd;
  var numberOfColumns = ss.getMaxColumns();
  

    
  // DCD 11/09/2020: Added logging  
  writeLog(JSON.stringify(e));
  
 // DCD 11/09/2020: Added logging
  writeLog("Editing sheet: " + e.source.getActiveSheet().getName() + " || rowStart: " + rowStart + " || rowEnd:  " + rowEnd +  " || cellStart: " + columnStart + " || cellEnd: " + columnEnd);
   
 // DCD 11/09/2020: Added logging
 // Number of Columns  
  writeLog("Number of Columns: " + numberOfColumns);
  
  
  // Are we modifying the status? 
  // DCD 11/10/2020: Check all the time regardless of what cell is changed
  //if(e.source.getActiveSheet().getName() === "Census" && currentCol === 15 && currentRow != 1)
  if(e.source.getActiveSheet().getName() === "Census")
  { 
    
      // Set column locations
      var statusCol = "O";
      var totalCol = "AE";
      var dateCol = "AD";    
      var reportCol = "AC";    
      var oldValCol = "AB";
      var oldRptDateCol = "AA";
            
       
    
     /// DCD 11/11/2020: Check to see if columns changed
     for( var columnNum = 1; columnNum < ss.getMaxColumns(); columnNum ++)
     {
       
       
       var currentColumn = columnToLetter(columnNum);        
       var colValue = ss.getRange(1, columnNum).getValue()
       
       
               
       if( ss.getRange(1, columnNum).getValue() === "Status")
       {
         statusCol = currentColumn;
       }
       else if( ss.getRange(1, columnNum).getValue() === "Previous Rpt Dt")
       {
         oldRptDateCol = currentColumn; 
       }else if( ss.getRange(1, columnNum).getValue() === "Previous Status")
       {
         oldValCol = currentColumn; 
       }else if( ss.getRange(1, columnNum).getValue() === "Report Start Date")
       {
         reportCol = currentColumn; 
       }else if( ss.getRange(1, columnNum).getValue() === "Last Updated")
       {
         dateCol = currentColumn; 
       }else if( ss.getRange(1, columnNum).getValue() === "Total Count")
       {
         totalCol = currentColumn; 
       }
   }    
      
    Logger.log("Status Column Set To: " + statusCol);
    Logger.log("Previous Rpt Dt Column Set To: " + oldRptDateCol);
    Logger.log("Previous Status Column Set To: " + oldValCol);
    Logger.log("Report Start Date Column Set To: " + reportCol);
    Logger.log("Last Update Date Column Set To: " + dateCol);
    Logger.log("Total Count Column Set To: " + totalCol);
    
    
    for( var currentRow = rowStart; currentRow <= rowEnd; currentRow ++)
    {    
    
      
      Logger.log("============================ Row: " + currentRow + " Start =============================" );
 
      
      var oldUpdateDate = "";
      var today = new Date().toISOString().slice(0,10);    
      var currentDate = new Date(); 
      
      
      // DCD 11/06/2020: Added logic for initiating
      if(ss.getRange(dateCol+currentRow).getValue() === "")      
      {       
        oldUpdateDate = "01/01/2001";
        
        // DCD 11/09/2020: Added logging
        writeLog("Old Update Date: NULL: Setting to: " + oldUpdateDate, currentRow);
        
      }else
      {
        oldUpdateDate = ss.getRange(dateCol+currentRow).getValue().toISOString().slice(0,10);
        
        // DCD 11/09/2020: Added logging
        writeLog("Old Update Date: " + oldUpdateDate, currentRow);
        
      }
      
      
      // Load current hour of the day and day of the week
      var hourOfDay = currentDate.getHours();
      
      // Load day of the week    
      var dayOfWeek = currentDate.getDay(); 
      
      // Get previous value of the drop down
      var prevVal = ss.getRange(statusCol+currentRow).getValue(); 
      if( currentCol === 15)
      {
        prevVal = e.oldValue;
      }
      
      
      // Get previous value of the previous value
      var prevPrevVal = ss.getRange(oldValCol+currentRow).getValue();
      
      // Week of Month
      var todayDate = new Date(today);
      var weekOfMonthToday = getWeekOfMonth( todayDate);
      var lastUpdateDate = new Date(oldUpdateDate);
      var weekOfMonthLastUpdateDate = getWeekOfMonth(lastUpdateDate);
      
      
      
      // If the compare date and the last update date the same? If not, ensure that we have not made a change for this week.
      var updateOk = 1;
      var loggingMesageState = "ok";    
      if(today === oldUpdateDate)
      {
        updateOk = 0;      
        loggingMesageState = "datefail";
      }else{
        if(weekOfMonthToday === weekOfMonthLastUpdateDate ) 
        {
          updateOk = 0;
          loggingMesageState = "weekfail";
        }
      }
      
      

      // DCD 11/09/2020: Added logging
      writeLog("           Today Info: Week of Month: " + weekOfMonthToday + " | Date: " + todayDate, currentRow );
      writeLog("Last Update Date Info: Week of Month: " + weekOfMonthLastUpdateDate + " | Date: " + lastUpdateDate, currentRow);
      
      if(loggingMesageState === "ok")
      {
        writeLog("Update Check: OK", currentRow);
      }else if (loggingMesageState === "datefail")
      {
        writeLog("Update Check: Failed because today and last update date are equal", currentRow);
      }else if(loggingMesageState === "weekfail")
      {
        writeLog("Update Check: Failed because last updated month and current month are the same.", currentRow);
      }
      
      
      
      // Set and format report date
      var rptDate = new Date(currentDate.setDate(currentDate.getDate() +5 - dayOfWeek)).toISOString().slice(0,10);
      
      // Sunday through Friday, Saturday's are set with the following logic
      if(dayOfWeek > 5)
      {
        rptDate = new Date(currentDate.setDate(currentDate.getDate() + dayOfWeek)).toISOString().slice(0,10);       
      }       
      

      /// DCD 11/10/2020: Final check, see if no value existed before that could mean that the row was pasted in
      if(prevVal === undefined)
      {
        writeLog("Because previous value was empty, we are forcing calculations. preVal: " + prevVal, currentRow);
        updateOk = 1;
      }else if(ss.getRange(totalCol+currentRow).getValue() === "")
      {
        writeLog("The total count value is set to empty either manually or some other mean. Forcing refresh: ", currentRow);
        updateOk = 1;
      }      
      
      
      /// DCD 11/10/2020: Final check for totalColumn
      writeLog("Status: " + ss.getRange(oldValCol+currentRow).getValue() + " | PrevStat: " + ss.getRange(statusCol+currentRow).getValue(), currentRow);
      if(ss.getRange(oldValCol+currentRow).getValue() === "Positive" || ss.getRange(statusCol+currentRow).getValue() === "Positive")
      {
        writeLog("Either the status or previous status was set to positive. Setting total count to 1.", currentRow);
        ss.getRange(totalCol+currentRow).setValue("1");                
      }else
      {
        writeLog("Neither status nor previous status = positive, resetting total count to 0", currentRow);
        ss.getRange(totalCol+currentRow).setValue("0");
      }
      
      
      // Set the oldValue if we have not changed it already for today
      if(updateOk === 1)
      { 
        // Set the updatedate date to today
        ss.getRange(dateCol+currentRow).setValue(today);      
        
        // Set the report col
        ss.getRange(reportCol+currentRow).setValue(rptDate);     
        
        // Calculate the old reportdate
        let dateParts = rptDate.split('-');
        var setOldRptDate = new Date(dateParts[0], dateParts[1]-1, dateParts[2]);
        
        // Just to make things more confusing, it turns out that the users want the report to report up to 3:00 on Fridays (not 12:00)
        // This logic address will set any changes made on Fridays after 3:00 to report the following week
        if(dayOfWeek == 5 && hourOfDay >= 15)
        {        
          var after3rptDate = new Date(dateParts[0], dateParts[1]-1, dateParts[2]);
          after3rptDate = new Date(after3rptDate.setDate(after3rptDate.getDate() + 7)).toISOString().slice(0,10);
          
          writeLog('Friday after 3PM Logic: Original RptDate: ' + rptDate + " |New RptDate: " + after3rptDate + " |Update Ok: " + updateOk , currentRow);
          
          ss.getRange(reportCol+currentRow).setValue(after3rptDate);     
        }
        
        
        var oldRptDate = new Date(setOldRptDate.setDate(setOldRptDate.getDate() - 7)).toISOString().slice(0,10);
        
        // Set previous value and old report date
        //writeLog("Starting Update");
        ss.getRange(oldValCol+currentRow).setValue(prevVal);                
        //writeLog("Ending Update");
        ss.getRange(oldRptDateCol+currentRow).setValue(oldRptDate); 
        //writeLog(updateOk === 1 );
        
        
        // This is a bit of a cluge to tricking GDS to work correctly        
        if(e.value === "Positive")
        {
          ss.getRange(oldRptDateCol+currentRow).setValue("01/01/2020");
          ss.getRange(totalCol + currentRow).setValue(1);
        }       
        
      }else
      {
        // Even though we end up here, the spreadsheet continues to update the 
        writeLog("No update needed, resetting previous-previous value: " + prevPrevVal, currentRow);
        ss.getRange(oldValCol+currentRow).setValue(prevPrevVal);
      }    
      
    
           
      
      
      Logger.log("============================ Row: " + currentRow + " Completed =============================" );
      
    } // end main for loop
    
  } // end main if
  
}



// Function name: getWeekOfMonth
//        Author: D.DiCesare - copied from outside source
//   Description: Calculates the week of a month
//    Parameters: Date: Date to extract week of month
//       Returns: Int: week number
//     Revisions: DCD 10/07/2020: Initial

function getWeekOfMonth(date) {  
 let adjustedDate = date.getDate()+date.getDay();
 let prefixes = ['0', '1', '2', '3', '4', '5'];
 return (parseInt(prefixes[0 | adjustedDate / 7])+1);
}




// Function name: writeLog
//        Author: D.DiCesare 
//   Description: Writes message to log and rownumber
//    Parameters: message: string of what to write into log
//                rownumber: row being processed
//       Returns: Int: NA
//     Revisions: DCD 11/10/2020: Initial
function writeLog(message, rowNumber)
{
  Logger.log( "Row: " + rowNumber + " : " + message);
}




// Function name: columnToLetter
//        Author: D.DiCesare 
//   Description: Converts number column to letter
//    Parameters: column: column number
//       Returns: column letter
//     Revisions: DCD 11/11/2020: Initial
function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}