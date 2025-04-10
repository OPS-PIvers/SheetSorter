/**
 * Utilities for Form Submissions Sorter
 * 
 * Helper functions to support the main script functionality.
 */

/**
 * Gets all unique values from a specific column in a sheet
 * @param {Sheet} sheet - The sheet to analyze
 * @param {number} columnIndex - The column index to get unique values from
 * @return {Array} Array of unique values
 */
function getUniqueValuesInColumn(sheet, columnIndex) {
  // Skip if sheet has only a header row or less
  if (sheet.getLastRow() <= 1) {
    return [];
  }
  
  // Get all values in the column (excluding header)
  const values = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1).getValues();
  
  // Extract to a flat array and filter out empty values
  const flatValues = values.flat().filter(value => value !== "");
  
  // Get unique values using Set
  return [...new Set(flatValues)];
}

/**
 * Batch processes rows to improve performance
 * @param {Sheet} sourceSheet - Source sheet with form responses
 * @param {number} sortColumnIndex - Index of column to sort by
 * @param {number} batchSize - Number of rows to process in each batch
 */
function batchProcessRows(sourceSheet, sortColumnIndex, batchSize = 20) {
  // Get configuration and check it exists
  const scriptProperties = PropertiesService.getScriptProperties();
  let processedRows = JSON.parse(scriptProperties.getProperty(PROPERTIES_KEYS.PROCESSED_ROWS) || "[]");
  
  // Get the data range (excluding header row)
  const lastRow = sourceSheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }
  
  // Process in batches to avoid timeout
  const totalBatches = Math.ceil((lastRow - 1) / batchSize);
  
  for (let batch = 0; batch < totalBatches; batch++) {
    const startRow = 2 + (batch * batchSize);
    const endRow = Math.min(startRow + batchSize - 1, lastRow);
    
    for (let row = startRow; row <= endRow; row++) {
      processRow(row, sortColumnIndex, sourceSheet);
    }
    
    // If this isn't the last batch, add a small delay to avoid hitting quota limits
    if (batch < totalBatches - 1) {
      Utilities.sleep(500);
    }
  }
}

/**
 * Cleans up the spreadsheet by removing empty sheets
 * @param {Spreadsheet} ss - The spreadsheet to clean
 * @return {number} Number of sheets removed
 */
function removeEmptySheets(ss) {
  const sheets = ss.getSheets();
  let removedCount = 0;
  
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    
    // Skip the form response sheet
    const config = getConfiguration();
    if (config && sheet.getSheetId() === parseInt(PropertiesService.getScriptProperties().getProperty(PROPERTIES_KEYS.FORM_SHEET_ID))) {
      continue;
    }
    
    // Check if sheet is empty (just has a header row)
    if (sheet.getLastRow() <= 1) {
      ss.deleteSheet(sheet);
      removedCount++;
    }
  }
  
  return removedCount;
}

/**
 * Export configuration to a JSON string
 * @return {string} JSON string of configuration
 */
function exportConfiguration() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const config = {
    sortColumnIndex: scriptProperties.getProperty(PROPERTIES_KEYS.SORT_COLUMN),
    formSheetId: scriptProperties.getProperty(PROPERTIES_KEYS.FORM_SHEET_ID),
    // Don't include processedRows to keep the export small
  };
  
  return JSON.stringify(config);
}

/**
 * Import configuration from a JSON string
 * @param {string} jsonConfig - JSON string of configuration
 * @return {boolean} Success state
 */
function importConfiguration(jsonConfig) {
  try {
    const config = JSON.parse(jsonConfig);
    const scriptProperties = PropertiesService.getScriptProperties();
    
    if (config.sortColumnIndex) {
      scriptProperties.setProperty(PROPERTIES_KEYS.SORT_COLUMN, config.sortColumnIndex);
    }
    
    if (config.formSheetId) {
      scriptProperties.setProperty(PROPERTIES_KEYS.FORM_SHEET_ID, config.formSheetId);
    }
    
    // Initialize processed rows if needed
    if (!scriptProperties.getProperty(PROPERTIES_KEYS.PROCESSED_ROWS)) {
      scriptProperties.setProperty(PROPERTIES_KEYS.PROCESSED_ROWS, JSON.stringify([]));
    }
    
    return true;
  } catch (error) {
    console.error("Failed to import configuration:", error);
    return false;
  }
}

/**
 * Creates all necessary sheets for existing unique values in the sort column
 * @param {Sheet} sourceSheet - Source sheet with form responses
 * @param {number} sortColumnIndex - Index of column to sort by
 * @return {number} Number of sheets created
 */
function createSheetsForUniqueValues(sourceSheet, sortColumnIndex) {
  // Get unique values
  const uniqueValues = getUniqueValuesInColumn(sourceSheet, sortColumnIndex);
  
  // Create sheets for each unique value
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let createdCount = 0;
  
  uniqueValues.forEach(value => {
    if (!value) return; // Skip empty values
    
    const sheetName = createValidSheetName(value.toString());
    
    // Check if sheet already exists
    if (!ss.getSheetByName(sheetName)) {
      // Create sheet and copy header
      const newSheet = ss.insertSheet(sheetName);
      
      // Copy the header row from the source sheet
      const headerRange = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn());
      const headerValues = headerRange.getValues();
      newSheet.getRange(1, 1, 1, headerValues[0].length).setValues(headerValues);
      
      // Format the header
      const headerFormat = headerRange.getTextStyles();
      const headerBackground = headerRange.getBackgrounds();
      newSheet.getRange(1, 1, 1, headerValues[0].length).setTextStyles(headerFormat);
      newSheet.getRange(1, 1, 1, headerValues[0].length).setBackgrounds(headerBackground);
      
      createdCount++;
    }
  });
  
  return createdCount;
}

/**
 * Reset the processed rows tracking to force re-processing of all rows
 */
function resetProcessedRowsTracking() {
  PropertiesService.getScriptProperties().setProperty(PROPERTIES_KEYS.PROCESSED_ROWS, JSON.stringify([]));
}

/**
 * Check if a value would create a valid sheet name
 * @param {string} value - Value to check
 * @return {boolean} Whether the value would create a valid sheet name
 */
function isValidSheetNameValue(value) {
  if (!value) return false;
  
  // Check length
  if (value.toString().length > 100) return false;
  
  // Check for invalid characters
  const invalidChars = /[\[\]\*\?\/\\:]/;
  return !invalidChars.test(value);
}

/**
 * Get statistics about the current sorting setup
 * @return {Object} Statistics object
 */
function getSortingStats() {
  try {
    const config = getConfiguration();
    if (!config) {
      return { error: "Not configured" };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const sourceSheet = config.sheet;
    const sortColumnIndex = config.sortColumnIndex;
    
    // Get unique values
    const uniqueValues = getUniqueValuesInColumn(sourceSheet, sortColumnIndex);
    
    // Count sorted sheets (excluding form sheet)
    let sortedSheets = 0;
    sheets.forEach(sheet => {
      if (sheet.getSheetId() !== sourceSheet.getSheetId()) {
        sortedSheets++;
      }
    });
    
    // Get processed rows count
    const processedRows = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROPERTIES_KEYS.PROCESSED_ROWS) || "[]");
    
    return {
      totalSubmissions: sourceSheet.getLastRow() - 1,
      uniqueCategories: uniqueValues.length,
      sortedSheets: sortedSheets,
      processedRows: processedRows.length,
      columnName: sourceSheet.getRange(1, sortColumnIndex).getValue()
    };
  } catch (error) {
    return { error: error.message };
  }
}

/**
 * Force update the header formatting on all sheets
 */
function updateAllSheetHeaders() {
  const config = getConfiguration();
  if (!config) return;
  
  const sourceSheet = config.sheet;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // Get header formatting from source sheet
  const headerRange = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn());
  const headerValues = headerRange.getValues();
  const headerFormat = headerRange.getTextStyles();
  const headerBackground = headerRange.getBackgrounds();
  
  // Update each sheet header (except source sheet)
  sheets.forEach(sheet => {
    if (sheet.getSheetId() !== sourceSheet.getSheetId()) {
      const lastCol = Math.min(sheet.getLastColumn(), headerValues[0].length);
      if (lastCol > 0) {
        sheet.getRange(1, 1, 1, lastCol).setValues(headerValues.map(row => row.slice(0, lastCol)));
        sheet.getRange(1, 1, 1, lastCol).setTextStyles(headerFormat.map(row => row.slice(0, lastCol)));
        sheet.getRange(1, 1, 1, lastCol).setBackgrounds(headerBackground.map(row => row.slice(0, lastCol)));
      }
    }
  });
}

/**
 * Generate a report of the current sorting setup
 * @return {string} HTML report content
 */
function generateSortingReport() {
  try {
    const config = getConfiguration();
    if (!config) {
      return "<p>The sorter has not been configured yet.</p>";
    }
    
    const stats = getSortingStats();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let html = `
      <h3>Form Submissions Sorter Report</h3>
      <p><strong>Spreadsheet:</strong> ${ss.getName()}</p>
      <p><strong>Source Sheet:</strong> ${config.sheet.getName()}</p>
      <p><strong>Sorting by:</strong> ${stats.columnName} (column ${config.sortColumnIndex})</p>
      <p><strong>Total Submissions:</strong> ${stats.totalSubmissions}</p>
      <p><strong>Unique Categories:</strong> ${stats.uniqueCategories}</p>
      <p><strong>Sorted Sheets:</strong> ${stats.sortedSheets}</p>
      <hr>
      <h4>Category Sheets:</h4>
      <ul>
    `;
    
    // List all sheets except the source sheet
    const sheets = ss.getSheets();
    sheets.forEach(sheet => {
      if (sheet.getSheetId() !== config.sheet.getSheetId()) {
        const rowCount = sheet.getLastRow() - 1; // Subtract header
        html += `<li>${sheet.getName()} (${rowCount} rows)</li>`;
      }
    });
    
    html += "</ul>";
    return html;
    
  } catch (error) {
    return `<p>Error generating report: ${error.message}</p>`;
  }
}
