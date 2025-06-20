/** @OnlyCurrentDoc */


function UpdateTemplate(){



      var now = new Date();
   
      var dayOrNight = Utilities.formatDate(now,"PST", "a");
      var pstTime = Utilities.formatDate(now,"PST", "h");
        
    //start to analyze if it starts at 2am
    if ((pstTime >= 2 && dayOrNight == "AM")){

      if (pstTime<=4){
        CopyNameTickerToTemplate() //copy-pastes the created ticker and stock names onto the template 
        PlotFormulasToTemplate()   //plots formulas to template
      }

      if (pstTime <5){
        UpdateRSI()                //updates the RSI and volume columns
      }

    }
}

function UpdateTemplateNotRSI(){

    
      
      // do stuff
      CopyNameTickerToTemplate() //copy-pastes the created ticker and stock names onto the template 
      PlotFormulasToTemplate()   //plots formulas to template

}

function UpdateTemplate2(){

     
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Template');
    
    // get the current loop counter
    var userProperties = PropertiesService.getUserProperties();
    var loopCounter = Number(userProperties.getProperty('loopCounter'));
    
    // put some limit on the number of loops
    // could be based on a calculation or user input
    // using a static number in this example
    var limit = 12;
    
    // if loop counter < limit number, run the repeatable action
    if (loopCounter < limit) {
      
      // see what the counter value is at the start of the loop
      Logger.log(loopCounter);
      
      // do stuff
      CopyNameTickerToTemplate() //copy-pastes the created ticker and stock names onto the template 
      PlotFormulasToTemplate()   //plots formulas to template
      UpdateRSI()                //updates the RSI and volume columns
      
      // increment the properties service counter for the loop
      loopCounter +=1;
      userProperties.setProperty('loopCounter', loopCounter);
      
      // see what the counter value is at the end of the loop
      Logger.log(loopCounter);
    }
    
    // if the loop counter is no longer smaller than the limit number
    // run this finishing code instead of the repeatable action block
    else {
      // Log message to confirm loop is finished
      //sheet.getRange(sheet.getLastRow()+1,1).setValue("Finished");
      Logger.log("Finished");
      
      // delete trigger because we've reached the end of the loop
      // this will end the program
      deleteTrigger();  
    }




}

function ExtractWillshire() {

//this function pulls data from the website into the extractor sheet based on the link source links

  var spreadsheet = SpreadsheetApp.getActive();
  var extractorSheet = spreadsheet.getSheetByName("Extract Data");
  var linkSheet = spreadsheet.getSheetByName("Link Source");
  var extractorLastRow = extractorSheet.getLastRow();
  var extractorLastCol = extractorSheet.getLastColumn();
  var linkLastRow = linkSheet.getLastRow();

  if (extractorLastRow>1){
    extractorSheet.getRange(1,1,extractorLastRow,extractorLastCol).clearContent();
  }

  extractorLastRow = 1;

  //importhtml('Link Source'!C32,"table",1)
for (var j = 1; j <= linkLastRow; j++) {     

     //plot import html formula to extractor sheet
    var extractorPlotRng = extractorSheet.getRange(extractorLastRow,1);
    extractorPlotRng.setFormula("=importhtml(" + "'Link Source'" + "!C" + j + ',"table",1)');

    

  extractorLastRow = extractorSheet.getLastRow() + 1; 
  SpreadsheetApp.flush();


}

extractorSheet.getRange('A' + 1 + ':I' + extractorLastRow).copyTo(extractorSheet.getRange('A' + 1 + ':I' + extractorLastRow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

GetNameAndTicker();

}

function GetNameAndTicker(){
  
  //this function creates the ticker and stock names of wilshire index stocks to be copy-pasted to the template

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = spreadsheet.getSheetByName("Extract Data");
 

  var startRow =findCellRowQuickExit("");
  var startCol = findCellColumnQuickExit("");
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  
  //sheet.getRange(startRow, 3, lastRow, lastCol).clearContent();
  var lastCol = sheet.getLastColumn();

var tickerCell = sheet.getRange(1,10);
//tickerCell.setFormula('=MID(B1,FIND("(",B1)+1,FIND(")",B1)-FIND("(",B1)-1)');
tickerCell.setFormula('=IFERROR(MID(A1,FIND("(",A1)+1,FIND(")*",A1)-FIND("(",A1)-1),"ERROR")');

var symbolCell = sheet.getRange(1,11);
//symbolCell.setFormula('=trim(MID(B1,1,FIND("(",B1)-2))');
symbolCell.setFormula('=trim(MID(A1,FIND("*",A1)+1,(FIND(" (",A1)-1)-FIND("*",A1)+1))');

var tickerRng = sheet.getRange(2,10,lastRow-1, 1);
var symbolRng = sheet.getRange(2,11, lastRow-1, 1);


tickerCell.copyTo(tickerRng);
 symbolCell.copyTo(symbolRng);

//delete errors after extract
filterThenDeleteErrors();

}

function CopyNameTickerToTemplate(){

  //this function copy-pastes the created ticker and stock names onto the template 


  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  
  var sheet = spreadsheet.getSheetByName("Extract Data");
  var templateSheet = spreadsheet.getSheetByName("Template");
  var lastRow = sheet.getLastRow();
  
  var tickerNsymbolRng = sheet.getRange(1,10,lastRow, 2);
  var tempTickerNsymbolRng = templateSheet.getRange(2,1,lastRow-1, 2);
  
tickerNsymbolRng.copyValuesToRange(templateSheet,1,2, 2,lastRow+1 );
 
}

function PlotFormulasToTemplate (){

  //this function plots the current price, 20-day SMA, 50-day SMA, RSI, Last Volume, and SDVA

  

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  
   
  var templateSheet = spreadsheet.getSheetByName("Template");
  templateSheet.activate;

  var lastRow = templateSheet.getLastRow();
  var realLastRow = lastRow;


      var linkSheet = spreadsheet.getSheetByName("Link Source");
      var now = new Date();
      var dayNow = Utilities.formatDate(now,"PST", "d");
      //var dayNow = 19
      var updatedDay = linkSheet.getRange("J3").getValue();

      var lastColm = templateSheet.getLastColumn();
    
 if (dayNow!=updatedDay){
        templateSheet.getRange(2,3,lastRow-1,11).clearContent();

        var lastRow = templateSheet.getLastRow();
        var realLastRow = lastRow;
        var lastColm = templateSheet.getLastColumn();
        var dayNowCell = linkSheet.getRange("J2");
        var updateDayCell = linkSheet.getRange("J3");
        dayNowCell.setValue(dayNow);
        updateDayCell.setValue(dayNow);

  }

var Avals = templateSheet.getRange("C1:C").getValues();
var Alast = Avals.filter(String).length+1;
var tempLastRow = Alast + 250;


if (tempLastRow <= lastRow){
  lastRow = tempLastRow
} 

if (Alast == realLastRow+1){
  return;
}

//Logger.log(lastRow);
//Logger.log(tempLastRow);

  //update current price
  var currPriceCell = templateSheet.getRange(Alast,3);
  currPriceCell.setFormula("=IFERROR(GOOGLEFINANCE(A" + Alast + "),0)");

var currPriceRng = templateSheet.getRange(Alast,3, 249, 1);

currPriceCell.copyTo(currPriceRng);

//update 20-Day SMA
  var fiftySMACell = templateSheet.getRange(Alast,4);
  fiftySMACell.setFormula('=IFERROR(AVERAGE(INDEX(GoogleFinance(A'+ Alast + ',"all",WORKDAY(TODAY(),-20),TODAY()),,5)),0)')

var fiftySMArng = templateSheet.getRange(Alast,4, 249, 1);

fiftySMACell.copyTo(fiftySMArng)

//update percent change of 20-day SMA
var fiftySMAchngCell = templateSheet.getRange(Alast,5);
  fiftySMAchngCell.setFormula("=IFERROR(((C" + Alast + "-D" + Alast + ")/D" + Alast +  "),0)" );

var fiftySMAchngRng = templateSheet.getRange(Alast,5, 249, 1)

fiftySMAchngCell.copyTo(fiftySMAchngRng)

//update 50-Day SMA
  var twohundSMACell = templateSheet.getRange(Alast,6);
  twohundSMACell.setFormula('=IFERROR(AVERAGE(INDEX(GoogleFinance(A' + Alast + ',"all",WORKDAY(TODAY(),-50),TODAY()),,5)),0)');

var twohundRng = templateSheet.getRange(Alast,6, 249, 1);

twohundSMACell.copyTo(twohundRng);

//update 14 day SDVA
  var sdvaCell = templateSheet.getRange(Alast,9);
  sdvaCell.setFormula('=IFERROR(AVERAGE(INDEX(GoogleFinance(A' + Alast + ',"all",WORKDAY(TODAY(),-14),TODAY()),,6)),0)');

var sdvaRng = templateSheet.getRange(Alast,9, 249, 1);

sdvaCell.copyTo(sdvaRng);


//=GOOGLEFINANCE(A2,"PE")
//=GOOGLEFINANCE(A2,"marketcap")

var peCell = templateSheet.getRange(Alast,12);
  peCell.setFormula('=IFERROR(GoogleFinance(A' + Alast + ',"PE"),0)');

var peRng = templateSheet.getRange(Alast,12, 249, 1);

peCell.copyTo(peRng);

var mkapCell = templateSheet.getRange(Alast,13);
  mkapCell.setFormula('=IFERROR(GoogleFinance(A' + Alast + ',"marketcap"),0)');

var mkapRng = templateSheet.getRange(Alast,13, 249, 1);

mkapCell.copyTo(mkapRng);

//=IF(and(AVERAGE(INDEX(GoogleFinance(A1101,"all",WORKDAY(TODAY(),-20),WORKDAY(TODAY(),-1)),,5))>AVERAGE(INDEX(GoogleFinance(A1101,"all",WORKDAY(TODAY(),-50),WORKDAY(TODAY(),-1)),,5)),D1101<F1101),"Death Cross",If(and(AVERAGE(INDEX(GoogleFinance(A1101,"all",WORKDAY(TODAY(),-20),WORKDAY(TODAY(),-1)),,5))<AVERAGE(INDEX(GoogleFinance(A1101,"all",WORKDAY(TODAY(),-50),WORKDAY(TODAY(),-1)),,5)),D1101>F1101),"Golden Cross",""))

//update cross formula
var crossCell = templateSheet.getRange(Alast,11);
var AlastPlus1 = Alast + 1;
   crossCell.setFormula('=IFERROR(IF(and(' + 'AVERAGE(INDEX(GoogleFinance(A' + Alast + ',"all",WORKDAY(TODAY(),-20),WORKDAY(TODAY(),-5)),,5))' + '>' + 'AVERAGE(INDEX(GoogleFinance(A' + Alast + ',"all",WORKDAY(TODAY(),-50),WORKDAY(TODAY(),-5)),,5))' +  ',D' + Alast + '<F' + Alast + '),"Death Cross",If(and(' + 'AVERAGE(INDEX(GoogleFinance(A' + Alast + ',"all",WORKDAY(TODAY(),-20),WORKDAY(TODAY(),-5)),,5))' + '<' + 'AVERAGE(INDEX(GoogleFinance(A' + Alast + ',"all",WORKDAY(TODAY(),-50),WORKDAY(TODAY(),-5)),,5))' +  ',D' + Alast + '>F' + Alast + '),"Golden Cross","n/a")),' + '"n/a"' + ")");

var crossRng = templateSheet.getRange(Alast,11, 249, 1);

crossCell.copyTo(crossRng);

if (lastRow == realLastRow){
templateSheet.getRange(lastRow,11).setValue("");
}

SpreadsheetApp.flush();
Utilities.sleep(5000);
//if (lastRow >= (realLastRow/4)){
templateSheet.getRange('A2:M' + lastRow).copyTo(templateSheet.getRange('A2:M' + lastRow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
//}

}


function UpdateRSI(){

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var templateSheet = spreadsheet.getSheetByName("Template");
  templateSheet.activate;

  var lastRow = templateSheet.getRange("C1:C").getValues().filter(String).length;
var realLastRow = lastRow;
var Avals = templateSheet.getRange("G1:G").getValues();
var Alast = Avals.filter(String).length+1;
var tempLastRow = Alast + 250;

var currPriceStartVal = templateSheet.getRange("C2").getValue();

if (currPriceStartVal==""){
  return;
}

if (tempLastRow <= lastRow){
  lastRow = tempLastRow
} 

if (Alast == realLastRow+1){
  SpreadsheetApp.flush();
    templateSheet.getRange('G' + 2 + ':J' + lastRow).copyTo(templateSheet.getRange('G' + 2 + ':J' + lastRow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  //filterThenDeleteNA
  return;
}


var MILLIS_PER_DAY = 1500 * 60 * 3;
var now = new Date();
var fourMilliLater = new Date(now.getTime() + MILLIS_PER_DAY);

//Logger.log(now)
//Logger.log(fourMilliLater)

//return;

//update 14-Day Volume average
var volAvgCell = templateSheet.getRange(Alast,10);
//var volAvgCell = templateSheet.getRange(2,10);
  volAvgCell.setFormula("=IFERROR(((H" + Alast + "-I" + Alast + ")/I" + Alast + "),0)");

var volAvgRng = templateSheet.getRange(Alast,10, 249, 1);
//var volAvgRng = templateSheet.getRange(2,10, lastRow-1, 1);

volAvgCell.copyTo(volAvgRng);

//update RSI
 var draftDataSheet =spreadsheet.getSheetByName("Data Draft");


var draftDataFormulaCell = draftDataSheet.getRange(1,1);

//update change formula
var draftDataChngCell = draftDataSheet.getRange(2,7);
draftDataSheet.getRange(1,7).setValue("Change");

//update RSI
var draftDataRSICell = draftDataSheet.getRange(2,8);
draftDataSheet.getRange(1,8).setValue("RSI");        


  for (var j = Alast; j <= lastRow; j++) {     


  var nowFromLoop = new Date();

  if (nowFromLoop >= fourMilliLater){
    SpreadsheetApp.flush();
    templateSheet.getRange('G' + 2 + ':J' + lastRow).copyTo(templateSheet.getRange('G' + 2 + ':J' + lastRow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    //filterThenDeleteNA
    return;
  }



     //update draft formula
      draftDataFormulaCell.setFormula('=sort(GoogleFinance(Template!A' + j + ',"all",WORKDAY(TODAY(),-50),TODAY()),1,false)');

var draftDataLastRow =draftDataSheet.getLastRow();

//update change formula
draftDataChngCell.setFormula("=E2-E3");
var draftDataChngRng = draftDataSheet.getRange(2,7, draftDataLastRow-1, 1);
draftDataChngCell.copyTo(draftDataChngRng);

//update RSI
    draftDataRSICell.setFormula('=IFERROR(100-(100/(1+((AVERAGEIF($G$2:$G$15,">0",$G$2:$G$15))/(-1*AVERAGEIF($G$2:$G$15,"<0",$G$2:$G$15))))),0)');
       SpreadsheetApp.flush();
      //plot to template the RSI
      var rsiCell = templateSheet.getRange(j,7);
        rsiCell.setFormula("='Data Draft'!H"+2);
      //plot to template yesterday volume
      var yestVolCell = templateSheet.getRange(j,8);
      //=if('Data Draft'!E1="","","")
      //=if('Data Draft'!F2=""0,'Data Draft'!F2)  
        yestVolCell.setFormula("=if('Data Draft'!F"+2+"="+'""' + ",0,"+ "'Data Draft'!F" +2 + ")");

        if (rsiCell.getValue=="#DIV/0!"){
          rsiCell.setValue('0');
        }

if (yestVolCell.getValue==""){
          yestVolCell.setValue('0');
        }

 SpreadsheetApp.flush();
      templateSheet.getRange('G' + j + ':H' + j).copyTo(templateSheet.getRange('G' + j + ':H' + j), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  

    }    

if (lastRow >= realLastRow){
    SpreadsheetApp.flush();
    templateSheet.getRange('G' + 2 + ':J' + lastRow).copyTo(templateSheet.getRange('G' + 2 + ':J' + lastRow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

}


}

function findCellRow(strKeyword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] == strKeyword) {
        row = values[i][j+1];
        //Logger.log(row);
       //Logger.log(i+1); // This is your row number
       return i+1;
      }
    }    
  }  
}

function findCellColumn(strKeyword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] == strKeyword) {
        row = values[i][j+1];
        //Logger.log(row);
       //Logger.log(j+1); // This is your row number
       return j+1;
      }
    }    
  }  
}


function findCellRowQuickExit(strKeyword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

rowLoop:
  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] != "" ) {
        row = values[i][j+1];
        //Logger.log(row);
       //Logger.log(i+1); // This is your row number
       return i+1;
       break rowLoop;
      }
    }    
  }  
}

function findCellColumnQuickExit(strKeyword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

rowLoop:
  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] != "" ) {
        row = values[i][j+1];
        //Logger.log(row);
       //Logger.log(j+1); // This is your row number
       return j+1;
       break rowLoop;
      }
    }    
  }  
}

function CopyPasteCell() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D2').activate();
  spreadsheet.getRange('D2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
};


function filterThenDeleteErrors() {

  //this function deletes rows with errors when data is pulled-up from website

  var spreadsheet = SpreadsheetApp.getActive();
  var extractSheet= spreadsheet.getSheetByName("Extract Data");
  var lastRow = extractSheet.getLastRow();

  //extractSheet.getFilter().remove();

  extractSheet.getRange('A1:K' + lastRow).createFilter();
  
  var criteria = SpreadsheetApp.newFilterCriteria()
  .whenTextContains('ERROR')
  .build();
  extractSheet.getFilter().setColumnFilterCriteria(10, criteria);
 
  extractSheet.deleteRows(1, lastRow);

}

function filterThenDeleteNA() {

  //this function deletes rows with errors when data is pulled-up from website

  //var spreadsheet = SpreadsheetApp.getActive();
  //var extractSheet= spreadsheet.getSheetByName("Template");
 // var lastRow = extractSheet.getLastRow();

//  var rangecells = sheet.getRange('#N/A');
 // rangecells.setValue('0')

  //var rangecells = sheet.getRange('#DIV/0');
  //rangecells.setValue('0')

//var range = extractSheet.getRange("A2:M" + lastRow);
//cell.setValue('#N/A','0');
//cell.setValue('#DIV/0','0');
//cell.setValue(' ','0');

 const sheet = SpreadsheetApp.getActiveSheet();
  //const range = sheet.getDataRange();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A2:M" + lastRow);

  const textFinder = range.createTextFinder('#N/A');
  const allOccurrences = textFinder.replaceAllWith('0');

const textFinder2 = range.createTextFinder('#DIV/0!');
  const allOccurrences2 = textFinder2.replaceAllWith('0');

//const textFinder3 = range.createTextFinder('');
  //const allOccurrences3 = textFinder3.replaceAllWith('0');

}

function onOpen() { 
  //var ui = SpreadsheetApp.getUi();
  
  //ui.createMenu("Auto Trigger")
    //.addItem("Run","runAuto")
    //.addToUi();
}

function runAuto() {
  
  // resets the loop counter if it's not 0
  refreshUserProps();
  
  // create trigger to run program automatically
  createTrigger();
}

function refreshUserProps() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('loopCounter', 0);
}

function createTrigger() {
  
  // Trigger every 1 minute
  ScriptApp.newTrigger('UpdateTemplate')
      .timeBased()
      .everyMinutes(5)
      .create();
}

function deleteTrigger() {
  
  // Loop over all triggers and delete them
  var allTriggers = ScriptApp.getProjectTriggers();
  
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}


function ReplaceDiv0() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
};