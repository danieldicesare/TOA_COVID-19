// Function name: onEdit
//        Author: D.DiCesare
//   Description: Event that fires off any time something happens on the sheets. Conditions placed to specficially look at the Census tab and the 
//                status column. 
//                *** Please note, this is a workable soltuion that ties into a report that displays on the internet (See Wally for report name and location). 
//                While this works, it is fragile given the nature of spreadsheets. It is advised that if the spreadsheet is to be used long term, an alternative
//                solution should be explored
//    Parameters: e: spreadsheet event
//       Returns: Int: NA
//     Revisions: 10/07/2020: Initial
function onEdit(e) {

  // Load column and range information
  var range = e.range; 
  var currentRow = range.getRow();
  var currentCol = range.getColumn();
  var ss = SpreadsheetApp.getActiveSheet();
  
  
  // Are we modifying the status?  
  if(e.source.getActiveSheet().getName() === "Census" && currentCol === 15 && currentRow != 1){    
       
    var dateCol = "AD";    
    var reportCol = "AC";    
    var oldValCol = "AB";
    var oldRptDateCol = "AA";
    var oldUpdateDate = ss.getRange(dateCol+currentRow).getValue().toISOString().slice(0,10);
    var today = new Date().toISOString().slice(0,10);    
    var currentDate = new Date();      
    
    // Load day of the week    
    var dayOfWeek = currentDate.getDay(); 
    
    // Get previous value of the drop down
    var prevVal = e.oldValue;
    
    // Get previous value of the previous value
    var prevPrevVal = ss.getRange(oldValCol+currentRow).getValue();
    
    // Week of Month
    var todayDate = new Date(today);
    var weekOfMonthToday = getWeekOfMonth( todayDate);
    var lastUpdateDate = new Date(oldUpdateDate);
    var weekOfMonthLastUpdateDate = getWeekOfMonth(lastUpdateDate);
    Logger.log("Week of Month This Week: " + weekOfMonthToday + " | Date: " + todayDate +  " || Week of Month Last Update: " + weekOfMonthLastUpdateDate + " | Date: " + lastUpdateDate);
    

    
    // Is the compare date and the last update date the same? If not, ensure that we have not made a change for this week.
    var updateOk = 1;
    if(today === oldUpdateDate)
    {
      updateOk = 0;
    }else{
      if(weekOfMonthToday === weekOfMonthLastUpdateDate ) 
      {
        updateOk = 0;
      }
    }
     
    // Set and format report date
    var rptDate = new Date(currentDate.setDate(currentDate.getDate() +5 - dayOfWeek)).toISOString().slice(0,10);
    
    // Sunday through Friday, Saturday's are set with the following logic
    if(dayOfWeek > 5)
    {
      rptDate = new Date(currentDate.setDate(currentDate.getDate() + dayOfWeek)).toISOString().slice(0,10);       
    }       
        
    Logger.log('OldReportUpdateDate: ' + oldUpdateDate + " |Today: " + today + " |Update Ok: " + updateOk );
    
   
    
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
      
      //ss.getRange("S"+currentRow).setValue(setOldRptDate);
      
      var oldRptDate = new Date(setOldRptDate.setDate(setOldRptDate.getDate() - 7)).toISOString().slice(0,10);
      
      // Set previous value and old report date
      //Logger.log("Starting Update");
      ss.getRange(oldValCol+currentRow).setValue(prevVal);                
      //Logger.log("Ending Update");
      ss.getRange(oldRptDateCol+currentRow).setValue(oldRptDate); 
      //Logger.log(updateOk === 1 );
      
      
      // This is a bit of a cluge to tricking GDS to work correctly        
      if(e.value === "Positive")
      {
        ss.getRange(oldRptDateCol+currentRow).setValue("01/01/2020");
      }
      
    }else
    {
      // Even though we end up here, the spreadsheet continues to update the 
      Logger.log("No update needed, resetting previous-previous value: " + prevPrevVal );
      ss.getRange(oldValCol+currentRow).setValue(prevPrevVal);
    }
    
  }
  
}



// Function name: getWeekOfMonth
//        Author: D.DiCesare - copied from outside source
//   Description: Calculates the week of a month
//    Parameters: Date: Date to extract week of month
//       Returns: Int: week number
//     Revisions: 10/07/2020: Initial

function getWeekOfMonth(date) {  
 let adjustedDate = date.getDate()+date.getDay();
 let prefixes = ['0', '1', '2', '3', '4', '5'];
 return (parseInt(prefixes[0 | adjustedDate / 7])+1);
}