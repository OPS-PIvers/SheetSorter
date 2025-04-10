/**
 * Form Submissions Sorter
 * 
 * This script automatically sorts Google Form submissions into separate sheets
 * based on values in a user-selected column.
 */

// Constants
const SCRIPT_TITLE = "Form Submissions Sorter";
const PROPERTIES_KEYS = {
  SORT_COLUMN: "sortColumnIndex",
  FORM_SHEET_ID: "formSheetId",
  PROCESSED_ROWS: "processedRows"
};

/**
 * Adds a custom menu to the spreadsheet when opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(SCRIPT_TITLE)
    .addItem("Setup Sorter", "showSetupDialog")
    .addItem("Sort Existing Submissions", "sortExistingSubmissions")
    .addItem("View Current Configuration", "showCurrentConfig")
    .addSeparator()
    .addItem("Reset Configuration", "resetConfiguration")
    .addToUi();
}

/**
 * Shows the setup dialog for the user to configure sorting options
 */
function showSetupDialog() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Check if this is a form response sheet
  if (!isFormResponseSheet(sheet)) {
    ui.alert(
      "Not a form response sheet",
      "Please select a sheet that's connected to a Google Form. This sheet should contain form responses.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Get the headers from the active sheet
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Create and show a list dialog for the user to select a column
  const htmlOutput = HtmlService.createHtmlOutput(
    createColumnSelectionHtml(headers)
  )
    .setWidth(400)
    .setHeight(300)
    .setTitle("Setup Form Submissions Sorter");
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Select Column for Sorting");
}

/**
 * Creates HTML for the column selection dialog
 * @param {Array} headers - Array of column headers
 * @return {string} HTML content
 */
function createColumnSelectionHtml(headers) {
  let html = '<p>Select the column you want to use for sorting submissions into separate sheets:</p>';
  html += '<select id="columnSelect" style="width: 100%; margin-bottom: 20px;">';
  
  headers.forEach((header, index) => {
    if (header) { // Only include non-empty headers
      html += `<option value="${index + 1}">${header}</option>`;
    }
  });
  
  html += '</select>';
  html += '<div style="text-align: center">';
  html += '<button onclick="saveSelection()">Save</button>';
  html += '</div>';
  
  // Add JavaScript to handle the selection
  html += `
    <script>
      function saveSelection() {
        const select = document.getElementById('columnSelect');
        const columnIndex = select.value;
        const columnName = select.options[select.selectedIndex].text;
        google.script.run
          .withSuccessHandler(closeDialog)
          .saveColumnSelection(columnIndex, columnName);
      }
      
      function closeDialog() {
        google.script.host.close();
      }
    </script>
  `;
  
  return html;
}

/**
 * Saves the selected column and sets up the trigger
 * @param {number} columnIndex - Index of the selected column
 * @param {string} columnName - Name of the selected column
 */
function saveColumnSelection(columnIndex, columnName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  try {
    // Save configuration to script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty(PROPERTIES_KEYS.SORT_COLUMN, columnIndex);
    scriptProperties.setProperty(PROPERTIES_KEYS.FORM_SHEET_ID, sheet.getSheetId().toString());
    
    // Create a tracking property for processed rows (empty array initially)
    if (!scriptProperties.getProperty(PROPERTIES_KEYS.PROCESSED_ROWS)) {
      scriptProperties.setProperty(PROPERTIES_KEYS.PROCESSED_ROWS, JSON.stringify([]));
    }
    
    // Set up trigger for form submissions
    setupFormSubmitTrigger();
    
    // Show success message
    ui.alert(
      "Setup Complete",
      `Form submissions will now be sorted based on the "${columnName}" column.\n\n` + 
      "A trigger has been set up to automatically sort new submissions.\n\n" +
      "You can also sort existing submissions from the menu.",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert("Error", `Setup failed: ${error.message}`, ui.ButtonSet.OK);
    console.error("Setup failed:", error);
  }
}

/**
 * Sets up a trigger for form submissions
 */
function setupFormSubmitTrigger() {
  // Delete any existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create a new trigger
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}

/**
 * Handles new form submissions
 * @param {Object} e - Form submit event object
 */
function onFormSubmit(e) {
  try {
    // Get the row number of the new submission
    const row = e.range.getRow();
    
    // Get configuration
    const config = getConfiguration();
    if (!config) {
      console.error("Configuration not found. Please run setup first.");
      return;
    }
    
    // Process the new submission
    processRow(row, config.sortColumnIndex, config.sheet);
    
  } catch (error) {
    console.error("Error processing form submission:", error);
  }
}

/**
 * Sorts all existing submissions
 */
function sortExistingSubmissions() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Get configuration
    const config = getConfiguration();
    if (!config) {
      ui.alert(
        "Configuration Not Found",
        "Please run the setup first to select which column to sort by.",
        ui.ButtonSet.OK
      );
      return;
    }
    
    const sheet = config.sheet;
    
    // Get the data range (excluding header row)
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      ui.alert("No Data", "No data found to sort.", ui.ButtonSet.OK);
      return;
    }
    
    // Process each row
    let processedCount = 0;
    for (let row = 2; row <= lastRow; row++) {
      const wasProcessed = processRow(row, config.sortColumnIndex, sheet);
      if (wasProcessed) processedCount++;
    }
    
    ui.alert(
      "Sorting Complete",
      `Processed ${processedCount} submissions.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert("Error", `Sorting failed: ${error.message}`, ui.ButtonSet.OK);
    console.error("Sorting failed:", error);
  }
}

/**
 * Process a single row and copy it to the appropriate sheet
 * @param {number} row - Row number to process
 * @param {number} sortColumnIndex - Index of the column to sort by
 * @param {Sheet} sourceSheet - Sheet containing the form responses
 * @return {boolean} Whether the row was newly processed
 */
function processRow(row, sortColumnIndex, sourceSheet) {
  // Get the tracking data
  const scriptProperties = PropertiesService.getScriptProperties();
  let processedRows = JSON.parse(scriptProperties.getProperty(PROPERTIES_KEYS.PROCESSED_ROWS) || "[]");
  
  // Skip if already processed (for manual sorting)
  const rowId = `${sourceSheet.getSheetId()}_${row}`;
  if (processedRows.includes(rowId)) {
    return false;
  }
  
  // Get the value in the sort column
  const sortValue = sourceSheet.getRange(row, sortColumnIndex).getValue();
  
  // Skip if no sort value
  if (!sortValue) {
    return false;
  }
  
  // Create a valid sheet name from the sort value
  const sheetName = createValidSheetName(sortValue.toString());
  
  // Get or create the destination sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let destSheet = ss.getSheetByName(sheetName);
  
  if (!destSheet) {
    // Create a new sheet with the sort value as the name
    destSheet = ss.insertSheet(sheetName);
    
    // Copy the header row from the source sheet
    const headerRange = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn());
    const headerValues = headerRange.getValues();
    destSheet.getRange(1, 1, 1, headerValues[0].length).setValues(headerValues);
    
    // Format the header
    const headerFormat = headerRange.getTextStyles();
    const headerBackground = headerRange.getBackgrounds();
    destSheet.getRange(1, 1, 1, headerValues[0].length).setTextStyles(headerFormat);
    destSheet.getRange(1, 1, 1, headerValues[0].length).setBackgrounds(headerBackground);
  }
  
  // Get the row data
  const rowRange = sourceSheet.getRange(row, 1, 1, sourceSheet.getLastColumn());
  const rowValues = rowRange.getValues();
  
  // Copy to the destination sheet (in the next available row)
  const nextRow = destSheet.getLastRow() + 1;
  destSheet.getRange(nextRow, 1, 1, rowValues[0].length).setValues(rowValues);
  
  // Copy formatting
  const rowFontStyles = rowRange.getTextStyles();
  const rowBackgrounds = rowRange.getBackgrounds();
  destSheet.getRange(nextRow, 1, 1, rowValues[0].length).setTextStyles(rowFontStyles);
  destSheet.getRange(nextRow, 1, 1, rowValues[0].length).setBackgrounds(rowBackgrounds);
  
  // Mark as processed
  processedRows.push(rowId);
  scriptProperties.setProperty(PROPERTIES_KEYS.PROCESSED_ROWS, JSON.stringify(processedRows));
  
  return true;
}

/**
 * Creates a valid sheet name from a given string
 * @param {string} name - Original name
 * @return {string} Valid sheet name
 */
function createValidSheetName(name) {
  // Remove invalid characters and trim
  let safeName = name.replace(/[\[\]\*\?\/\\]/g, " ").trim();
  
  // Ensure the name isn't too long (max 100 chars for sheet names)
  if (safeName.length > 100) {
    safeName = safeName.substring(0, 100);
  }
  
  // Ensure the name isn't empty
  if (!safeName) {
    safeName = "Unnamed";
  }
  
  return safeName;
}

/**
 * Shows the current configuration
 */
function showCurrentConfig() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfiguration();
  
  if (!config) {
    ui.alert(
      "No Configuration",
      "The sorter has not been set up yet. Please run the setup first.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Get the column name
  const columnName = config.sheet.getRange(1, config.sortColumnIndex).getValue();
  
  ui.alert(
    "Current Configuration",
    `Sorting submissions based on: ${columnName}\n` +
    `Sheet: ${config.sheet.getName()}\n` +
    `Column: "${columnName}" (column ${config.sortColumnIndex})`,
    ui.ButtonSet.OK
  );
}

/**
 * Resets the configuration
 */
function resetConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Reset Configuration",
    "Are you sure you want to reset the configuration? This will remove the trigger and all settings.",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // Delete triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Clear properties
    PropertiesService.getScriptProperties().deleteAllProperties();
    
    ui.alert(
      "Reset Complete",
      "The configuration has been reset. You can run the setup again when needed.",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert("Error", `Reset failed: ${error.message}`, ui.ButtonSet.OK);
    console.error("Reset failed:", error);
  }
}

/**
 * Gets the current configuration
 * @return {Object|null} Configuration object or null if not set up
 */
function getConfiguration() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sortColumnIndex = scriptProperties.getProperty(PROPERTIES_KEYS.SORT_COLUMN);
  const formSheetId = scriptProperties.getProperty(PROPERTIES_KEYS.FORM_SHEET_ID);
  
  if (!sortColumnIndex || !formSheetId) {
    return null;
  }
  
  // Find the sheet by ID
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let formSheet = null;
  
  // Loop through all sheets to find the one with matching ID
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId().toString() === formSheetId) {
      formSheet = sheets[i];
      break;
    }
  }
  
  if (!formSheet) {
    return null;
  }
  
  return {
    sortColumnIndex: parseInt(sortColumnIndex),
    sheet: formSheet
  };
}

/**
 * Checks if a sheet is connected to a form
 * @param {Sheet} sheet - Sheet to check
 * @return {boolean} Whether the sheet is a form response sheet
 */
function isFormResponseSheet(sheet) {
  try {
    // Method 1: Try to get the form URL directly
    const formUrl = sheet.getFormUrl();
    if (formUrl && formUrl.length > 0) {
      return true;
    }
    
    // Method 2: Check the sheet properties (sometimes works when getFormUrl fails)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formTriggers = ScriptApp.getUserTriggers(ss);
    for (let i = 0; i < formTriggers.length; i++) {
      if (formTriggers[i].getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
        // A form trigger exists for this spreadsheet
        return true;
      }
    }
    
    // Method 3: Check if the sheet has form item response headers
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.includes("Timestamp") && headers.includes("Email Address")) {
      // Likely a form response sheet based on common form headers
      return true;
    }
    
    return false;
  } catch (e) {
    console.error("Error checking if sheet is form response sheet:", e);
    return false;
  }
}
