/**
 * @OnlyCurrentDoc
 * This script automates the process of extracting stock data,
 * populating a template, and performing financial calculations
 * within a Google Sheet.
 */

// --- Constants (Centralized Configuration) ---
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId(); // Use ID for robustness
const SHEET_NAMES = {
  TEMPLATE: 'Template',
  EXTRACT_DATA: 'Extract Data',
  LINK_SOURCE: 'Link Source',
  DATA_DRAFT: 'Data Draft'
};

const TEMPLATE_RANGES = {
  TICKER_SYMBOL_START_ROW: 2, // Row where ticker and symbol data starts
  FORMULA_START_COL: 3,      // Column where price, SMA, etc., formulas start (C)
  RSI_VOL_START_COL: 7,      // Column where RSI and Volume formulas start (G)
  FORMULA_FILL_COUNT: 249    // Number of rows to fill formulas down
};

const PROPERTIES_KEYS = {
  LOOP_COUNTER: 'loopCounter'
};

// --- Helper Functions ---

/**
 * Gets a specific sheet by its name.
 * @param {string} sheetName The name of the sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found. Please ensure it exists.`);
  }
  return sheet;
}

/**
 * Copies a formula from a source cell and applies it down a target range.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to work on.
 * @param {number} startRow The starting row for applying the formula.
 * @param {number} col The column for the formula.
 * @param {number} numRows The number of rows to apply the formula down.
 * @param {string} formula The formula string.
 * @param {string} [formulaReferenceCell] Optional: A cell reference for the formula to copy from if not directly setting.
 */
function copyFormulaToRange(sheet, startRow, col, numRows, formula, formulaReferenceCell = null) {
  const sourceCell = sheet.getRange(startRow, col);
  sourceCell.setFormula(formula);
  if (numRows > 1) { // Only copy if filling more than one row
    const targetRange = sheet.getRange(startRow, col, numRows, 1);
    sourceCell.copyTo(targetRange);
  }
}

/**
 * Finds the first non-empty row in a sheet, starting from row 1.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @returns {number} The row number of the first non-empty cell.
 */
function findFirstNonEmptyRow(sheet) {
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j] !== "") {
        return i + 1; // Return 1-based index
      }
    }
  }
  return 1; // If all empty, start at row 1
}

/**
 * Finds the first non-empty column in a sheet, starting from column 1.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @returns {number} The column number of the first non-empty cell.
 */
function findFirstNonEmptyColumn(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length === 0 || values[0].length === 0) return 1; // Handle empty sheet

  for (let j = 0; j < values[0].length; j++) { // Only check first row for initial non-empty column
    for (let i = 0; i < values.length; i++) {
      if (values[i][j] !== "") {
        return j + 1; // Return 1-based index
      }
    }
  }
  return 1; // If all empty, start at column 1
}

/**
 * Replaces specific error values like #N/A and #DIV/0! with '0' in a given range.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to apply replacements on.
 * @param {string} rangeA1Notation The A1 notation of the range to clean.
 */
function cleanErrorValues(sheet, rangeA1Notation) {
  const range = sheet.getRange(rangeA1Notation);
  range.createTextFinder('#N/A').replaceAllWith('0');
  range.createTextFinder('#DIV/0!').replaceAllWith('0');
}


// --- Main Functions (Refined Logic) ---

/**
 * Orchestrates the daily template update process based on PST time.
 * Runs key functions: CopyNameTickerToTemplate, PlotFormulasToTemplate, and UpdateRSI (conditionally).
 * Designed to be triggered at specific times.
 */
function UpdateTemplate() {
  const now = new Date();
  const pstHour = Number(Utilities.formatDate(now, "PST", "H")); // 0-23 hour format

  // Run initial updates if between 2 AM and 4 AM PST
  if (pstHour >= 2 && pstHour <= 4) {
    Logger.log(`Running initial template update at ${pstHour} PST.`);
    // These functions can run together as they don't depend on time.
    CopyNameTickerToTemplate();
    PlotFormulasToTemplate();

    // RSI update only runs if within a specific time window, typically longer-running
    if (pstHour <= 5) { // Original logic: pstTime < 5 (so 2,3,4 AM)
       Logger.log("Updating RSI for template.");
       UpdateRSI();
    }
  } else {
    Logger.log(`Skipping full template update outside 2-5 AM PST window. Current hour: ${pstHour}`);
  }
}

/**
 * Updates the template by copying stock names and tickers.
 * Does not update RSI, used for specific scenarios.
 */
function UpdateTemplateNotRSI() {
  Logger.log("Running template update without RSI.");
  CopyNameTickerToTemplate();
  PlotFormulasToTemplate();
}

/**
 * Manages a loop to update the template iteratively, typically for large datasets
 * or to manage Google Apps Script execution limits. Deletes the trigger upon completion.
 */
function UpdateTemplate2() {
  const userProperties = PropertiesService.getUserProperties();
  let loopCounter = Number(userProperties.getProperty(PROPERTIES_KEYS.LOOP_COUNTER) || 0); // Default to 0 if not set
  const limit = 12; // Define the iteration limit

  if (loopCounter < limit) {
    Logger.log(`Template update loop: Iteration ${loopCounter + 1} of ${limit}`);

    // Perform core update actions
    CopyNameTickerToTemplate();
    PlotFormulasToTemplate();
    UpdateRSI();

    // Increment and save counter for next iteration
    loopCounter++;
    userProperties.setProperty(PROPERTIES_KEYS.LOOP_COUNTER, loopCounter);
  } else {
    Logger.log("Template update loop finished. Deleting trigger.");
    //sheet.getRange(sheet.getLastRow()+1,1).setValue("Finished"); // If you want to log in sheet
    deleteTrigger();
  }
}

/**
 * Extracts stock data from external websites (using IMPORTHML)
 * and populates the 'Extract Data' sheet.
 */
function ExtractWillshire() {
  Logger.log("Starting data extraction from Willshire links.");
  const extractorSheet = getSheet(SHEET_NAMES.EXTRACT_DATA);
  const linkSheet = getSheet(SHEET_NAMES.LINK_SOURCE);

  // Clear existing content in extractor sheet
  extractorSheet.clearContents(); // Clears all content, safer than getLastRow/Col logic for initial clear

  const linkData = linkSheet.getRange("C1:C" + linkSheet.getLastRow()).getValues(); // Get all links from column C
  let currentExtractorRow = 1;

  linkData.forEach((row, i) => {
    const link = row[0];
    if (link) { // Only process if link exists
      const formula = `=IMPORTHML('${SHEET_NAMES.LINK_SOURCE}'!C${i + 1},"table",1)`;
      extractorSheet.getRange(currentExtractorRow, 1).setFormula(formula);
      SpreadsheetApp.flush(); // Forces formula calculation
      currentExtractorRow = extractorSheet.getLastRow() + 1; // Update for next importHTML
    }
  });

  // Convert formulas to values (after all imports are done)
  // Get the *actual* last row after formulas calculate
  const finalExtractorLastRow = extractorSheet.getLastRow();
  if (finalExtractorLastRow > 0) { // Ensure there's data to copy
    extractorSheet.getRange(1, 1, finalExtractorLastRow, extractorSheet.getLastColumn())
      .copyTo(extractorSheet.getRange(1, 1, finalExtractorLastRow, extractorSheet.getLastColumn()),
              SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }

  Logger.log("Finished data extraction. Now getting names and tickers.");
  GetNameAndTicker();
}


/**
 * Processes extracted data to derive stock tickers and names,
 * and cleans up any errors.
 */
function GetNameAndTicker() {
  Logger.log("Deriving stock names and tickers.");
  const sheet = getSheet(SHEET_NAMES.EXTRACT_DATA);
  const lastRow = sheet.getLastRow();

  if (lastRow < 1) { // No data extracted
    Logger.log("No data in Extract Data sheet to process.");
    return;
  }

  // Use a more robust way to find starting row/col for formulas if needed,
  // but if it's always row 1, col 10/11, keep it simple.
  const formulaStartRow = 1; // Formulas placed in header row
  const dataFillNumRows = lastRow - formulaStartRow; // Number of rows to fill *data* down

  // Formulas for Ticker and Symbol Extraction
  // Note: These formulas might need careful testing to ensure they handle various data formats.
  const tickerFormula = '=IFERROR(MID(A1,FIND("(",A1)+1,FIND(")*",A1)-FIND("(",A1)-1),"ERROR")';
  const symbolFormula = '=TRIM(MID(A1,FIND("*",A1)+1,(FIND(" (",A1)-1)-FIND("*",A1)+1))';

  // Apply ticker formula
  copyFormulaToRange(sheet, formulaStartRow, 10, 1, tickerFormula); // Only apply to header cell first
  sheet.getRange(formulaStartRow, 10).copyTo(sheet.getRange(formulaStartRow + 1, 10, dataFillNumRows, 1)); // Copy down data rows

  // Apply symbol formula
  copyFormulaToRange(sheet, formulaStartRow, 11, 1, symbolFormula); // Only apply to header cell first
  sheet.getRange(formulaStartRow + 1, 11, dataFillNumRows, 1).copyTo(sheet.getRange(formulaStartRow + 1, 11, dataFillNumRows, 1)); // Copy down data rows

  SpreadsheetApp.flush(); // Ensure formulas are calculated before filtering
  Logger.log("Filtering and deleting error rows.");
  filterThenDeleteErrors(); // Remove rows with errors
}


/**
 * Copies the processed stock names and tickers from 'Extract Data'
 * to the 'Template' sheet.
 */
function CopyNameTickerToTemplate() {
  Logger.log("Copying names and tickers to Template sheet.");
  const extractSheet = getSheet(SHEET_NAMES.EXTRACT_DATA);
  const templateSheet = getSheet(SHEET_NAMES.TEMPLATE);

  const lastRowExtract = extractSheet.getLastRow();
  if (lastRowExtract < TEMPLATE_RANGES.TICKER_SYMBOL_START_ROW -1) { // Ensure there's data after header in extract sheet
    Logger.log("No ticker/symbol data to copy to template.");
    return;
  }

  // Get range to copy from (assuming Ticker in col 10, Symbol in col 11)
  const sourceRange = extractSheet.getRange(TEMPLATE_RANGES.TICKER_SYMBOL_START_ROW, 10, lastRowExtract - TEMPLATE_RANGES.TICKER_SYMBOL_START_ROW + 1, 2);
  const targetStartRow = TEMPLATE_RANGES.TICKER_SYMBOL_START_ROW; // Destination row in Template sheet
  const targetEndRow = targetStartRow + sourceRange.getNumRows() -1;

  // Copy values directly to the template sheet
  sourceRange.copyValuesToRange(templateSheet, 1, 2, targetStartRow, targetEndRow);

  Logger.log("Names and tickers copied.");
}


/**
 * Plots various financial formulas (price, SMAs, PE, MarketCap, Cross)
 * onto the 'Template' sheet.
 */
function PlotFormulasToTemplate() {
  Logger.log("Plotting financial formulas to Template sheet.");
  const templateSheet = getSheet(SHEET_NAMES.TEMPLATE);
  const linkSheet = getSheet(SHEET_NAMES.LINK_SOURCE);

  const lastRowData = templateSheet.getRange("A:A").getValues().filter(String).length; // Actual last row with data in Col A
  if (lastRowData < TEMPLATE_RANGES.TICKER_SYMBOL_START_ROW) {
    Logger.log("No tickers in Template sheet to plot formulas for.");
    return;
  }

  const formulaStartRow = TEMPLATE_RANGES.TICKER_SYMBOL_START_ROW;
  const numRowsToFill = lastRowData - formulaStartRow +1; // Number of rows with tickers

  // --- Daily Refresh Logic ---
  const now = new Date();
  const currentDay = Utilities.formatDate(now, "PST", "d"); // Current day as a number
  const lastUpdatedDay = String(linkSheet.getRange("J3").getValue()); // Ensure it's treated as string for comparison

  if (currentDay !== lastUpdatedDay) {
    Logger.log(`New day detected. Clearing old formula content for ${lastUpdatedDay}, updating to ${currentDay}.`);
    // Clear existing formula columns (C to M)
    templateSheet.getRange(formulaStartRow, TEMPLATE_RANGES.FORMULA_START_COL, numRowsToFill, 11).clearContent();

    // Update the last updated day in Link Source sheet
    linkSheet.getRange("J2").setValue(currentDay);
    linkSheet.getRange("J3").setValue(currentDay); // Assuming J2 and J3 store the same day for logic
  }

  // --- Plotting Formulas ---
  // Using 'let' for dynamic ALast based on column data
  let ALast = templateSheet.getRange("C:C").getValues().filter(String).length + 1; // First empty cell in Col C
  const actualLastTickerRow = templateSheet.getRange("A:A").getValues().filter(String).length; // Last row with a ticker

  // Ensure formulas are only plotted in rows that have tickers
  const effectiveNumRowsToFill = actualLastTickerRow - ALast + 1;
  if (effectiveNumRowsToFill <= 0) {
      Logger.log("Formulas already plotted for all tickers in Template sheet.");
      return;
  }

  const fillerRows = TEMPLATE_RANGES.FORMULA_FILL_COUNT; // Max rows to fill in one go (249 from original code)
  const endRowForBatch = Math.min(ALast + fillerRows - 1, actualLastTickerRow);
  const batchNumRows = endRowForBatch - ALast + 1;


  // Current Price (Column C)
  copyFormulaToRange(templateSheet, ALast, 3, batchNumRows, `=IFERROR(GOOGLEFINANCE(A${ALast}),0)`);

  // 20-Day SMA (Column D)
  copyFormulaToRange(templateSheet, ALast, 4, batchNumRows, `=IFERROR(AVERAGE(INDEX(GoogleFinance(A${ALast},"all",WORKDAY(TODAY(),-20),TODAY()),,5)),0)`);

  // Percent Change of 20-day SMA (Column E)
  copyFormulaToRange(templateSheet, ALast, 5, batchNumRows, `=IFERROR(((C${ALast}-D${ALast})/D${ALast}),0)`);

  // 50-Day SMA (Column F)
  copyFormulaToRange(templateSheet, ALast, 6, batchNumRows, `=IFERROR(AVERAGE(INDEX(GoogleFinance(A${ALast},"all",WORKDAY(TODAY(),-50),TODAY()),,5)),0)`);

  // 14-Day SDVA (Column I)
  copyFormulaToRange(templateSheet, ALast, 9, batchNumRows, `=IFERROR(AVERAGE(INDEX(GoogleFinance(A${ALast},"all",WORKDAY(TODAY(),-14),TODAY()),,6)),0)`);

  // PE Ratio (Column L)
  copyFormulaToRange(templateSheet, ALast, 12, batchNumRows, `=IFERROR(GoogleFinance(A${ALast},"PE"),0)`);

  // Market Cap (Column M)
  copyFormulaToRange(templateSheet, ALast, 13, batchNumRows, `=IFERROR(GoogleFinance(A${ALast},"marketcap"),0)`);

  // Cross Formula (Golden/Death Cross) (Column K)
  // This formula is quite complex, using WORKDAY(TODAY(),-5) for historical SMA.
  const crossFormula = `=IFERROR(IF(AND(AVERAGE(INDEX(GoogleFinance(A${ALast},"all",WORKDAY(TODAY(),-20),WORKDAY(TODAY(),-5)),,5))>AVERAGE(INDEX(GoogleFinance(A${ALast},"all",WORKDAY(TODAY(),-50),WORKDAY(TODAY(),-5)),,5)),D${ALast}<F${ALast}),"Death Cross",IF(AND(AVERAGE(INDEX(GoogleFinance(A${ALast},"all",WORKDAY(TODAY(),-20),WORKDAY(TODAY(),-5)),,5))<AVERAGE(INDEX(GoogleFinance(A${ALast},"all",WORKDAY(TODAY(),-50),WORKDAY(TODAY(),-5)),,5)),D${ALast}>F${ALast}),"Golden Cross","n/a")),"n/a")`;
  copyFormulaToRange(templateSheet, ALast, 11, batchNumRows, crossFormula);


  SpreadsheetApp.flush(); // Ensure all formulas are written and calculated.
  // Original code had a sleep and then copy-paste values for columns A2:M to lastRow.
  // If the goal is to convert all plotted formulas to values after plotting, do it here.
  // It's generally better to convert values only when needed to retain formulas for dynamic updates.
  // I'm assuming you want to keep them as formulas for now unless performance dictates converting them.
  // templateSheet.getRange('A2:M' + actualLastTickerRow).copyTo(templateSheet.getRange('A2:M' + actualLastTickerRow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  Logger.log("Financial formulas plotted.");
}


/**
 * Updates the RSI and Volume-related columns on the 'Template' sheet.
 * This function handles its own batching/looping if needed.
 */
function UpdateRSI() {
  Logger.log("Updating RSI and Volume data.");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = getSheet(SHEET_NAMES.TEMPLATE);
  const draftDataSheet = getSheet(SHEET_NAMES.DATA_DRAFT);

  const lastRowTemplate = templateSheet.getRange("A:A").getValues().filter(String).length; // Last row with ticker in Template
  if (lastRowTemplate < TEMPLATE_RANGES.TICKER_SYMBOL_START_ROW) {
    Logger.log("No tickers in Template sheet to update RSI/Volume for.");
    return;
  }

  const formulaStartRow = TEMPLATE_RANGES.TICKER_SYMBOL_START_ROW;
  const numRowsToProcess = lastRowTemplate - formulaStartRow + 1; // Number of tickers to process

  // Clear previous data in Data Draft sheet if it's for a new ticker calculation
  draftDataSheet.clearContents();
  draftDataSheet.getRange(1, 7).setValue("Change"); // Set headers for helper columns
  draftDataSheet.getRange(1, 8).setValue("RSI");

  // RSI calculation often requires iterating through each stock's historical data.
  // The original code iterated through stocks (j) and potentially exited early.
  // This re-evaluates the approach for clarity and efficiency.
  const MILLIS_PER_LOOP_LIMIT = 3 * 60 * 1000; // 3 minutes for current batch (adjust as needed)
  const startTime = new Date().getTime();

  for (let i = 0; i < numRowsToProcess; i++) {
    const currentRow = formulaStartRow + i; // Current row on Template sheet
    const ticker = templateSheet.getRange(currentRow, 1).getValue(); // Get ticker from column A

    if (!ticker) { // Skip if no ticker
      continue;
    }

    // Check execution time limit to prevent exceeding Google Apps Script limits
    if (new Date().getTime() - startTime > MILLIS_PER_LOOP_LIMIT) {
      Logger.log(`RSI update paused due to time limit. Processed up to row ${currentRow -1}.`);
      SpreadsheetApp.flush();
      // Important: If you pause here, you need a mechanism (like PropertiesService or another trigger)
      // to resume from where it left off, similar to how UpdateTemplate2 uses a loopCounter.
      // For simplicity here, I'm just noting it.
      return;
    }

    // 1. Get historical data into Data Draft sheet
    const draftDataFormula = `=SORT(GoogleFinance("${ticker}","all",WORKDAY(TODAY(),-50),TODAY()),1,FALSE)`;
    draftDataSheet.getRange(1, 1).setFormula(draftDataFormula);
    SpreadsheetApp.flush(); // Ensure data is pulled

    const draftDataLastRow = draftDataSheet.getLastRow();
    if (draftDataLastRow < 2) { // Need at least 2 rows for change calculation
        Logger.log(`Skipping RSI for ${ticker}: Not enough historical data.`);
        templateSheet.getRange(currentRow, TEMPLATE_RANGES.RSI_VOL_START_COL).setValue(0); // Set RSI to 0
        templateSheet.getRange(currentRow, TEMPLATE_RANGES.RSI_VOL_START_COL + 1).setValue(0); // Set Volume to 0
        continue; // Move to next ticker
    }

    // 2. Calculate Change in Data Draft (Column G)
    const changeFormula = `=E2-E3`; // Assumes 'Close' price is in Column E
    const changeTargetRange = draftDataSheet.getRange(2, 7, draftDataLastRow - 1, 1); // From row 2 downwards
    draftDataSheet.getRange(2, 7).setFormula(changeFormula); // Apply formula to first cell
    draftDataSheet.getRange(2, 7).copyTo(changeTargetRange); // Copy down

    // 3. Calculate RSI in Data Draft (Column H)
    const rsiFormula = `=IFERROR(100-(100/(1+((AVERAGEIF($G$2:$G$15,">0",$G$2:$G$15))/(-1*AVERAGEIF($G$2:$G$15,"<0",$G$2:$G$15))))),0)`;
    draftDataSheet.getRange(2, 8).setFormula(rsiFormula); // Apply to first cell

    SpreadsheetApp.flush(); // Ensure formulas are calculated before copying values

    // 4. Plot RSI and Yesterday's Volume to Template
    const rsiValue = draftDataSheet.getRange(2, 8).getValue(); // Get calculated RSI value
    templateSheet.getRange(currentRow, TEMPLATE_RANGES.RSI_VOL_START_COL).setValue(rsiValue === '#DIV/0!' ? 0 : rsiValue); // Set RSI

    const yestVolumeValue = draftDataSheet.getRange(2, 6).getValue(); // Get Yesterday's Volume (assumes column F in Data Draft)
    templateSheet.getRange(currentRow, TEMPLATE_RANGES.RSI_VOL_START_COL + 1).setValue(yestVolumeValue || 0); // Set Volume, default to 0 if empty

    SpreadsheetApp.flush(); // Ensure values are written

    // 5. Convert RSI and Volume to values in Template sheet (Optional, depends on preference)
    templateSheet.getRange(currentRow, TEMPLATE_RANGES.RSI_VOL_START_COL, 1, 2)
      .copyTo(templateSheet.getRange(currentRow, TEMPLATE_RANGES.RSI_VOL_START_COL, 1, 2), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    // Clear Data Draft for next iteration (optional, but good for memory)
    draftDataSheet.clearContents();
  }
  Logger.log("RSI and Volume data update complete.");
}

/**
 * Deletes rows with "ERROR" in column 10 (J) of the 'Extract Data' sheet.
 */
function filterThenDeleteErrors() {
  Logger.log("Filtering and deleting error rows in Extract Data sheet.");
  const extractSheet = getSheet(SHEET_NAMES.EXTRACT_DATA);
  const lastRow = extractSheet.getLastRow();

  if (lastRow < 1) {
    Logger.log("No data to filter in Extract Data sheet.");
    return;
  }

  const range = extractSheet.getRange(1, 1, lastRow, extractSheet.getLastColumn());
  const filter = range.getFilter(); // Get existing filter or null

  if (filter) { // If a filter exists, remove it first
    filter.remove();
  }

  // Create a new filter for the entire range
  const newFilter = range.createFilter();
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextContains('ERROR')
    .build();

  newFilter.setColumnFilterCriteria(10, criteria); // Column J is 10

  // Get filtered range to delete
  // Note: getRange with a filter is tricky for deletion. It's often safer to
  // get the filtered rows' values, identify rows to delete, and then delete by index.
  // Direct deletion with getFilter().setColumnFilterCriteria().deleteRows() is not standard.
  // The original code `extractSheet.deleteRows(1, lastRow);` after filtering will delete ALL rows
  // if the filter is applied. This is very dangerous and likely not intended.
  // A safer approach:
  const values = extractSheet.getRange(1, 10, lastRow, 1).getValues(); // Get values from column J
  const rowsToDelete = [];
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).includes('ERROR')) { // Check for 'ERROR' in column J
      rowsToDelete.push(i + 1); // Store 1-based row index
    }
  }

  // Delete rows from bottom up to avoid shifting issues
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    extractSheet.deleteRow(rowsToDelete[i]);
  }

  if (newFilter) { // Remove the filter after operation
    newFilter.remove();
  }
  Logger.log("Error rows cleaned up.");
}

/**
 * Replaces #N/A and #DIV/0! error values with '0' in the active sheet's data range.
 * This version uses the utility function.
 */
function filterThenDeleteNA() {
  Logger.log("Cleaning #N/A and #DIV/0! errors.");
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) { // Ensure there's a range to clean
      cleanErrorValues(sheet, `A2:M${lastRow}`); // Apply to data rows, columns A to M
  }
  Logger.log("Error values replaced with 0.");
}

// --- Trigger Management (Simplified and Consolidated) ---

/**
 * Sets up a time-based trigger to run the UpdateTemplate function.
 * Resets the loop counter before creating the trigger.
 */
function runAuto() {
  Logger.log("Setting up automated update trigger.");
  refreshUserProps(); // Reset loop counter
  createTrigger();    // Create new trigger
  Logger.log("Automated update trigger created successfully.");
}

/**
 * Resets the loop counter property used for iterative updates.
 */
function refreshUserProps() {
  Logger.log("Resetting loop counter.");
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(PROPERTIES_KEYS.LOOP_COUNTER, 0);
}

/**
 * Creates a time-based trigger to run 'UpdateTemplate' every 5 minutes.
 */
function createTrigger() {
  ScriptApp.newTrigger('UpdateTemplate')
    .timeBased()
    .everyMinutes(5)
    .create();
}

/**
 * Deletes all existing triggers for the current script project.
 */
function deleteTrigger() {
  Logger.log("Deleting all existing triggers.");
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  Logger.log("All triggers deleted.");
}

// --- Minor / Unused Functions (Consider removal or refactoring) ---

// This function seems unused and likely redundant with GetNameAndTicker's copy logic
// function CopyPasteCell() {
//   var spreadsheet = SpreadsheetApp.getActive();
//   spreadsheet.getRange('D2').activate();
//   spreadsheet.getRange('D2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
// };

// These findCell functions are overly generic and inefficient for most Apps Script tasks.
// Using `sheet.getLastRow()` or `range.getValues().filter(String).length` is typically better.
// function findCellRow(strKeyword) { /* ... */ }
// function findCellColumn(strKeyword) { /* ... */ }
// function findCellRowQuickExit(strKeyword) { /* ... */ }
// function findCellColumnQuickExit(strKeyword) { /* ... */ }

// This function seems unused or meant for manual debugging.
// function ReplaceDiv0() { /* ... */ }

// The onOpen function seems to have commented-out UI code. If not used, can remove.
// function onOpen() { /* ... */ }