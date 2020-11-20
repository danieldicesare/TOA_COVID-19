// Function name: onEdit
//        Author: D.DiCesare / W McKenzie
//   Description: Event that fires off any time something happens on the sheets. Conditions placed to specficially look at the Census tab and the 
//                status column. 
//                *** Please note, this is a workable soltuion that ties into a report that displays on the internet (See Wally for report name and location). 
//                While this works, it is fragile given the nature of spreadsheets. It is advised that if the spreadsheet is to be used long term, an alternative
//                solution should be explored
//    Parameters: e: spreadsheet event
//       Returns: Int: NA
//     Revisions: DCD 10/07/2020: Initial
//                DCD 10/16/2020: New requirement: Added Friday 3PM logic, voiced opinion that data should be real time
//                DCD 10/21/2020: New requirement: adjusted requirement to show all schools not just schools with COVID. Added Total Count column and populted
//                                                 Added rows 2 through 14 with base values to trick report into totalling correctly
//                                                 Added new tab Schools with name of school and sort order
//                DCD 11/09/2020: Added additional logging.
//                DCD 11/10/2020: Added logic for total count and e.prevVal
//                DCD 11/11/2020: New requirement: support coulmns shifting. Added logic to find the necessary columns for the report. 
//                DCD 11/17/2020: New requirement: support break out of schools into tabs and update "census" tab will all data
//                                                 finally, users now want data as of last night as opposed to last Friday, continue to voice that data should be real time 
//                DCD 11/20/2020: Final version designed to extract data from aggregate tab and store in columns on the agreggate tab which will drive the report on the andover site. It has been 
//                                discussed that the number of dependencies that this approach has is somewhat dangerous. However, the risks have been deemed to be acceptable. 
function onEdit(e) {

  if(e.source.getActiveSheet().getName() != "Aggreg_Dashboard"){
    snapShot();
  }
  
  
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

 
// Function name: getreportNextStartDate
//        Author: D.DiCesare 
//   Description: Derives next report start date 
//    Parameters: currentDate - date to use for starting point
//                dayOfWeekToStart - "Next, Prior, Mon, Tue, Wed, Thu, Fri, Sat, Sun"
//       Returns: next report start date 
//     Revisions: DCD 11/17/2020: Initial
function getreportNextStartDate(currentDate, dayOfWeekToStart)
{
  var reportNextStartDate = "";
  
  // Load current hour of the day and day of the week
  var hourOfDay = currentDate.getHours();
  
  // Load day of the week    
  var dayOfWeek = currentDate.getDay(); 
  
  var reportNextStartDate = "";
  
  // This logic is incomplete, we need to build case for each day 
  if(dayOfWeekToStart === "Next")
  {
     // Set and format report date and set the previous report day to the required prior reporting period
    var reportNextStartDate = new Date(currentDate.setDate(currentDate.getDate() + 1)).toISOString().slice(0,10);
  }else if(dayOfWeekToStart === "Prior")
  {
     // Set and format report date and set the previous report day to the required prior reporting period
    var reportNextStartDate = new Date(currentDate.setDate(currentDate.getDate() - 1)).toISOString().slice(0,10);
  }else if(dayOfWeekToStart === "Fri")
  {
    // Set and format report date and set the previous report day to the required prior reporting period
    var reportNextStartDate = new Date(currentDate.setDate(currentDate.getDate() +5 - dayOfWeek)).toISOString().slice(0,10);
    
    // Sunday through Friday, Saturday's are set with the following logic
    if(dayOfWeek > 5)
    {
      reportNextStartDate = new Date(currentDate.setDate(currentDate.getDate() + dayOfWeek)).toISOString().slice(0,10);       
    }       
  }
  
  
  return reportNextStartDate; 
}

 
// Function name: getReportPreviousStartDate
//        Author: D.DiCesare 
//   Description: Derives previous report start date 
//    Parameters: reportNextStartDate
//              : daysPrior
//       Returns: previous report start date 
//     Revisions: DCD 11/17/2020: Initial
function getReportPreviousStartDate(reportNextStartDate, daysPrior)
{
  let dateParts = reportNextStartDate.split('-');
  var setreportPreviousStartDate = new Date(dateParts[0], dateParts[1]-1, dateParts[2]);
  var reportPreviousStartDate = new Date(setreportPreviousStartDate.setDate(setreportPreviousStartDate.getDate() - daysPrior)).toISOString().slice(0,10);
  return reportPreviousStartDate;
}


// Function name: snapShot
//        Author: D.DiCesare 
//   Description: Gets snapshot of data as of last night and stores in columns on aggregate sheet.  
//    Parameters: NA
//       Returns: previous report start date 
//     Revisions: DCD 11/17/2020: Initial
function snapShot()
{
  Logger.log("SnapShot Method" );
  
  var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var aggTab = activeSpreadSheet.getSheetByName("Aggreg_Dashboard");  
  var tabName = aggTab.getName()
  var today = new Date().toISOString().slice(0,10);
  
  // Load last snapshot date info
  var lastSnapDate = aggTab.getRange(2,12).getValue().toISOString().slice(0,10);
  
  if( lastSnapDate < today)
  {    
    aggTab.getRange(2,12).setValue(today);
    
    var maxRow = aggTab.getMaxRows();

    Logger.log("MaxRows: " + maxRow);    
        
    for(var currentRow = 2; currentRow < maxRow; currentRow ++)
    {
      if(aggTab.getRange(currentRow,1).getValue() === "Total")
      {
        break;
      }
      
      Logger.log(aggTab.getRange(currentRow,1).getValue() + ": " + aggTab.getRange(currentRow,2).getValue());
      
      aggTab.getRange(currentRow,13).setValue(aggTab.getRange(currentRow,2).getValue());
      
    }
  }
  
  
  
  Logger.log(lastSnapDate);
  
  
}




