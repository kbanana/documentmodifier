/**
 * @file Defines operations that can be performed on a file.
 */


/**
 * Duplicates a tab for a spreadsheet.
 * @param {string} fileId - The file ID of the spreadsheet to modify.
 * @param {string} oldTabName - The name of the tab to duplicate.
 * @param {string} newTabName - The name to give the new, duplicated tab.
 * @returns {boolean} - true if the duplication succeeded, false if it failed.
 */
function duplicateTab(fileId, oldTabName, newTabName)
{
  var spreadsheet = SpreadsheetApp.openById(fileId);
  
  // Grab the 'oldTabName' sheet.
  var oldTab = spreadsheet.getSheetByName(oldTabName);
  if (oldTab == null)
  {
    Logger.log("Unable to locate tab '" + oldTabName + "' on spreadsheet '" + spreadsheet.getName() + "'.");
    return false;
  }
  
  // Delete the 'newTabName' sheet if it already exists.
  var existingTab = spreadsheet.getSheetByName(newTabName);
  if (existingTab != null)
  {
    spreadsheet.deleteSheet(existingTab);
  }
  
  // Duplicate the old sheet.
  oldTab.activate();
  var newTab = spreadsheet.duplicateActiveSheet();
  
  // Rename the new sheet to 'newTabName' and then hide it.
  newTab.setName(newTabName);
  newTab.hideSheet();
  
  return true;
}

/**
 * Wraps and centers text for a spreadsheet.
 * @param {string} fileId - The file ID of the spreadsheet to modify.
 * @returns {boolean} - true if the operation succeeded, false if it failed.
 */
function wrapAndCenterText(fileId)
{
  // Open the file in Google Spreadsheets.
  var spreadsheet = SpreadsheetApp.openById(fileId);
  
  // Retrive the list of sheets.
  var allSheets = spreadsheet.getSheets();
  
  // For each sheet in the document, loop over the cells to apply our changes.
  for (var sheetIndex in allSheets)
  {
    var currentSheet = allSheets[sheetIndex];
    
    // Loop over the relevant rows and grab the cells we care about.
    var rowArray = [23, 28, 33, 49, 69, 88];
    for (var rowIndex in rowArray)
    {
      // Grab the cell we care about.
      var cell = currentSheet.getRange('A' + rowArray[rowIndex]);
      
      // Set middle alignment for the cell.
      cell.setVerticalAlignment('middle');
      
      // Set text wrapping for the cell. 
      cell.setWrap(true);
    }
  }
  return true;
}

/**
 * Hides a spreadsheet tab.
 * @param {string} fileId - The file ID of the spreadsheet to modify.
 * @param {string} tabName - The name of the tab to hide.
 * @returns {boolean} - true if the operation succeeded, false if it failed.
 */
function hideTab(fileId, tabName)
{
  var spreadsheet = SpreadsheetApp.openById(fileId);
  
  // Grab the 'tabName' sheet.
  var tab = spreadsheet.getSheetByName(tabName);
  if (tab == null)
  {
    Logger.log("Unable to locate tab '" + tabName + "' on spreadsheet '" + spreadsheet.getName() + "'.");
    return false;
  }
  
  // Only proceed if the tab isn't already hidden.
  if (!tab.isSheetHidden())
  {
    // Count the visible tabs.
    var visibleCount = 0;
    var allSheets = spreadsheet.getSheets();
    for (var i in allSheets)
    {
      if (!allSheets[i].isSheetHidden())
      {
        visibleCount++;
      }
    }
    
    // Make sure there are at least two visible tabs, otherwise hiding one will fail.
    if (visibleCount == 1)
    {
      Logger.log("Unable to hide tab: spreadsheet '" + spreadsheet.getName() + "' only has one visible tab.");
      return false;
    }
    
    // Hide the tab.
    tab.hideSheet();
  }
  
  return true;
}

/**
 * Shows a spreadsheet tab.
 * @param {string} fileId - The file ID of the spreadsheet to modify.
 * @param {string} tabName - The name of the tab to show.
 * @returns {boolean} - true if the operation succeeded, false if it failed.
 */
function showTab(fileId, tabName)
{
  var spreadsheet = SpreadsheetApp.openById(fileId);
  
  // Grab the 'tabName' sheet.
  var tab = spreadsheet.getSheetByName(tabName);
  if (tab == null)
  {
    Logger.log("Unable to locate tab '" + tabName + "' on spreadsheet '" + spreadsheet.getName() + "'.");
    return false;
  }
  
  // Show the tab.
  tab.showSheet();
  
  return true;
}
