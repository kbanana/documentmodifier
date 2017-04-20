/**
 * @file Helper functions for converting spreadsheets to PDFs and downloading files to the user's Drive.
 */


/**
 * Helper function to save an arbitrary file from the Internet as a file in the user's Drive.
 * @param {string} url - The url of the file to download, e.g. http://example.com/files/foo.pdf
 * @param {string} folderName - The name of the folder in Drive where the file should be saved, or null to save in the root folder.
 * @param {string} fileName - The filename to use for the newly-created file.
 * @param {boolean} isInternalUrl - true if the file is being downloaded from the user's Google Docs or Sheets (requires an OAuth token to access).
 * @returns {boolean} - true if the download succeeded, false otherwise.
 */
function downloadFileToDrive(url, folderName, fileName, isInternalUrl)
{
  // Prepare the download options for our web request.
  var fetchOptions = null;
  if (isInternalUrl)
  {
    fetchOptions =
    {
      headers:
      {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    };
  }
  
  // Download the file at the specified URL, and abort if we encounter an error.
  var webResponse = UrlFetchApp.fetch(url, fetchOptions);
  if (Math.floor(webResponse.getResponseCode() / 100) != 2)
  {
    Logger.log("Unable to download file! Response code was " + webResponse.getResponseCode() + ".");
    return false;
  }
  
  // Determine which folder to use for storing the download.
  var folder;
  if (folderName == null)
  {
    // If no folderName was specified, just use the root Drive folder.
    folder = DriveApp.getRootFolder();
  }
  else
  {
    // Locate the destination folder in Google Drive (there should only be one folder with this name!)
    var folders = DriveApp.getFoldersByName(folderName);
    var folderMatches = [];
    while (folders.hasNext())
    {
      folderMatches.push(folders.next());
    }
    
    // Abort if we find multiple folders with the same name.
    if (folderMatches.length != 1)
    {
      Logger.log("Unable to find the destination folder in Drive. Found " + folderMatches.length + " matching folders, expected exactly 1.");
      return false;
    }
    
    folder = folderMatches[0];
  }
  
  // Abort if the destination folder already contains a file with the desired filename.
  if (folder.getFilesByName(fileName).hasNext())
  {
    Logger.log("Unable to upload the file to Drive - there's already a file named '" + fileName + "' in the destination folder.");
    return false;
  }
  
  // All preflight checks are complete. We can safely create the file and save it in Drive.
  var data = webResponse.getBlob();
  var newFile = DriveApp.createFile(data);
  newFile.setName(fileName);
  folder.addFile(newFile);
  
  Logger.log("File downloaded and saved to Drive!");
  return true;
}

/**
 * Helper function to save a spreadsheet as a PDF into the user's Drive.
 * @param {string} fileId - The File ID of the spreadsheet to save.
 * @param {string} tabName - The name of the spreadsheet tab that should be saved as a PDF.
 * @param {string} folderName - The name of the folder in Drive where the PDF should be saved, or null to save in the root folder.
 * @param {string} fileName - The filename to use for the newly-created PDF.
 * @param {boolean} isPortrait - Whether or not to export the PDF in portrait mode.
 * @returns {boolean} - true if the download succeeded, false otherwise.
 */
function saveSheetAsPDF(fileId, tabName, folderName, fileName, isPortrait)
{
  // Open the file in Google Spreadsheets.
  var spreadsheet = SpreadsheetApp.openById(fileId);
  
  // Grab the 'tabName' sheet.
  var tab = spreadsheet.getSheetByName(tabName);
  if (tab == null || tab.isSheetHidden())
  {
    Logger.log("Error: tab '" + tabName + "' on spreadsheet '" + spreadsheet.getName() + "' is either missing or invisible.");
    return false;
  }
  
  // Build the URL that requests a PDF conversion from Google Drive.
  var options = '&gid=' + tab.getSheetId()
  + '&size=letter'      // paper size
  + '&portrait=' + isPortrait    // orientation, false for landscape
  + '&fitw=true'        // fit to width, false for actual size
  + '&sheetnames=false&printtitle=false&pagenumbers=true'  //hide optional headers and footers
  + '&gridlines=true'  // show gridlines
  + '&fzr=false'; // do not repeat row headers (frozen rows) on each page
  
  // Build the full PDF url.
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?exportFormat=pdf&format=pdf' + options;
  
  // Download the URL to the user's Drive.
  Logger.log("Saving file as PDF...");
  return downloadFileToDrive(url, folderName, fileName, true);
}
