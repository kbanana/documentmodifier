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
