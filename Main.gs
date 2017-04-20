/**
 * @file Main toolkit entry point. These are the functions that we call directly to perform our operations.
 */


/**
 * Finds all spreadsheets in the 'Evals' folder, duplicates their 'FALL16' tabs to create 'SPRING17' tabs.
 */
function CreateSpringSheets()
{
  var fileIds = getFilesInFolder('Evals');
  for (var i in fileIds)
  {
    duplicateTab(fileIds[i], 'FALL16', 'SPRING17');
  }
}

/**
 * Wraps and centers certain cells for all spreadsheets in the 'Evals' folder.
 */
function WrapAndCenterInputCells()
{
  var fileIds = getFilesInFolder('Evals');
  for (var i in fileIds)
  {
    wrapAndCenterText(fileIds[i]);
  }
}

/**
 * Shows 'SPRING17' tabs and hides 'FALL16' tabs for all spreadsheets in the 'Evals' folder.
 */
function ShowSpringHideFall()
{
  var fileIds = getFilesInFolder('Evals');
  for (var i in fileIds)
  {
    showTab(fileIds[i], 'SPRING17');
    hideTab(fileIds[i], 'FALL16');
  }
}

/**
 * Downloads the Wikipedia logo file and stores it into the 'MyDownloads' folder.
 */
function DownloadWikipediaExample()
{
  downloadFileToDrive('https://www.wikipedia.org/portal/wikipedia.org/assets/img/Wikipedia-logo-v2.png', 'MyDownloads', 'wikipedia.png', false);
}

/**
 * Creates a PDF copy of the 'FALL16' tab of the 'SampleEval' spreadsheet and stores it into the 'MyDownloads' folder.
 */
function DownloadPdfExample()
{
  var matchingFiles = DriveApp.getFilesByName('SampleEval');
  if (matchingFiles.hasNext())
  {
    var evalFile = matchingFiles.next();
    saveSheetAsPDF(evalFile.getId(), 'FALL16', 'MyDownloads', 'Sample Eval PDF', true);
  }
}
