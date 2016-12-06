function createSpringSheets() {
  // Grab the root folder.
  var rootFolderIterator = DriveApp.getFoldersByName('Evals');
  if (rootFolderIterator.hasNext()) {
    var rootFolder = rootFolderIterator.next();
    
    // Step over all the subfolders in the root folder.
    var subfolderIterator = rootFolder.getFolders();
    while (subfolderIterator.hasNext()) {
      var subfolder = subfolderIterator.next();
      
      // Step over all the files in the current subfolder.
      var fileIterator = subfolder.getFiles();
      while (fileIterator.hasNext()) {
        var file = fileIterator.next();
        
        // Open the file in Google Spreadsheets.
        var spreadsheet = SpreadsheetApp.openById(file.getId());
        
        // Delete the 'SPRING' sheet if it already exists.
        var existingSpringSheet = spreadsheet.getSheetByName('SPRING2017');
        if (existingSpringSheet != null) {
        spreadsheet.deleteSheet(existingSpringSheet);
        }
        
        // Rename the active sheet to 'FALL16' (there should only be one sheet for now).
        spreadsheet.getActiveSheet().setName('FALL16');
        
        // Duplicate the active sheet.
        var newSheet = spreadsheet.duplicateActiveSheet();
        
        // Rename the new sheet to 'SPRING17' and then hide it.
        newSheet.setName('SPRING17');
        newSheet.hideSheet();
      }
    }
  }
}


function WrapandCenterText() {
  // Grab the root folder.
  var rootFolderIterator = DriveApp.getFoldersByName('Evals');
  if (rootFolderIterator.hasNext()) {
    var rootFolder = rootFolderIterator.next();
    
    // Step over all the subfolders in the root folder.
    var subfolderIterator = rootFolder.getFolders();
    while (subfolderIterator.hasNext()) {
      var subfolder = subfolderIterator.next();
      
      // Step over all the files in the current subfolder.
      var fileIterator = subfolder.getFiles();
      while (fileIterator.hasNext()) {
        var file = fileIterator.next();
        
        // Open the file in Google Spreadsheets.
        var spreadsheet = SpreadsheetApp.openById(file.getId());
        
        // Retrive the list of sheets.
        var allSheets = spreadsheet.getSheets();
        
        // For each sheet in the document, loop over the cells to apply our changes.
        for (var sheetIndex in allSheets) {
          var currentSheet = allSheets[sheetIndex];
        
          // Loop over the relevant rows and grab the cells we care about.
          var rowArray = [23, 28, 33, 49, 69, 88];
          for (var rowIndex in rowArray) {
            
            // Grab the cell we care about.
            var cell = currentSheet.getRange('A' + rowArray[rowIndex]);
            
            // Set middle alignment for the cell.
            cell.setVerticalAlignment('middle');
            
            // Set text wrapping for the cell. 
            cell.setWrap(true);
          }
        }
      }
    }  
  }
}



function saveSheetstoPDF() {
  // Grab the root folder.
  var rootFolderIterator = DriveApp.getFoldersByName('TestEvals');
  if (rootFolderIterator.hasNext()) {
    var rootFolder = rootFolderIterator.next();
    
    // Step over all the subfolders in the root folder.
    var subfolderIterator = rootFolder.getFolders();
    while (subfolderIterator.hasNext()) {
      var subfolder = subfolderIterator.next();
      
      // Step over all the files in the current subfolder.
      var fileIterator = subfolder.getFiles();
      while (fileIterator.hasNext()) {
        var file = fileIterator.next();
        
        // Open the file in Google Spreadsheets.
        var spreadsheet = SpreadsheetApp.openById(file.getId());
        
        // Export file to PDF
        var pdfexport = 'export?exportFormat=pdf&format=pdf'
        var options = '&gid=' + spreadsheet.getSheetByName('FALL16')
        + '&size=letter'      // paper size
        + '&portrait=true'    // orientation, false for landscape
        + '&fitw=true'        // fit to width, false for actual size
        + '&sheetnames=false&printtitle=false&pagenumbers=true'  //hide optional headers and footers
        + '&gridlines=true'  // show gridlines
        + '&fzr=false';       // do not repeat row headers (frozen rows) on each page
        
      }
    }  
  } 
} 

/**
 * Downloads a file from the internet and saves it to your Google Drive.
 * @param {string} foldername - The name of the folder in Drive where the file should be saved.
 * @param {string} filename - The filename to use for the newly-created file.
 * @param {string} url - The url of the file to download, e.g. http://example.com/files/foo.pdf
 */
function saveFileToDrive(foldername, filename, url) {
  // Download the file at the specified URL, and abort if we encounter an error.
  var webResponse = UrlFetchApp.fetch(url);
  if (Math.floor(webResponse.getResponseCode() / 100) != 2) {
    Logger.log("Unable to download file! Response code was " + webResponse.getResponseCode() + ".");
    return;
  }
  
  // Locate the destination folder in Google Drive (there should only be one folder with this name!)
  var folders = DriveApp.getFoldersByName(foldername);
  var folderMatches = [];
  while (folders.hasNext()) {
    folderMatches.push(folders.next());
  }
  
  // Abort if we find multiple folders with the same name.
  if (folderMatches.length != 1) {
    Logger.log("Unable to find the destination folder in Drive. Found " + folderMatches.length + " matching folders, expected exactly 1.");
    return;
  }
  
  // Abort if the destination folder already contains a file with the desired filename.
  var folder = folderMatches[0];
  if (folder.getFilesByName(filename).hasNext()) {
    Logger.log("Unable to upload the file to Drive - there's already a file named " + filename + " in the destination folder.");
    return;
  }
  
  // All preflight checks are complete. We can safely create the file and save it in Drive.
  var data = webResponse.getBlob();
  var newFile = DriveApp.createFile(data);
  newFile.setName(filename);
  folder.addFile(newFile);
  Logger.log("File downloaded and saved to Drive!");
}
