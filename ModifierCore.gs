/**
 * @file Helper functions for performing batch actions on many files at once.
 */


/**
 * Helper function to obtain all the file IDs inside a folder or nested folder structure.
 * @param {string} folderName - The name of the folder to search through.
 * @returns {array} - A list of all the file IDs in the folder or any subfolder.
 */
function getFilesInFolder(folderName)
{
  // Start by locating the root folder.
  var foldersToProcess = [];
  var folderIterator = DriveApp.getFoldersByName(folderName);
  while (folderIterator.hasNext())
  {
    foldersToProcess.push(folderIterator.next());
  }
  
  // Make sure we found exactly one folder with the target folderName.
  if (foldersToProcess.length != 1)
  {
    Logger.log("Unable to find the target folder in Drive. Found " + foldersToProcess.length + " matching folders, expected exactly 1.");
    return [];
  }
  
  // Recursively process all subfolders.
  var fileIds = [];
  while (foldersToProcess.length > 0)
  {
    // Pop a folder off the processing queue.
    var folder = foldersToProcess.pop();
  
    // Add each subfolder to the processing queue.
    var subfolderIterator = folder.getFolders();
    while (subfolderIterator.hasNext())
    {
      foldersToProcess.push(subfolderIterator.next());
    }
    
    // Add each file ID to the return array.
    var fileIterator = folder.getFiles();
    while (fileIterator.hasNext())
    {
      var file = fileIterator.next();
      fileIds.push(file.getId());
    }
  }
  
  return fileIds;
}
