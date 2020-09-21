function onOpen() {

  // Add a custom menu to run the script
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var HackerMan = [{name: "Mail To Drive", functionName: "saveGmailtoGoogleDrive"},{name: " ♛ UNZIP ♛", functionName: "unzipFiles"}];
  ss.addMenu(" ♛ ℍOMEFℝIEND DEVMOD ♛ ", HackerMan);
};

function saveGmailtoGoogleDrive() {
var endDate =  new Date(new Date().getTime() - (1 * 24 * 60 * 60 * 1000));
var realDate = Utilities.formatDate(endDate, "GMT+1", "dd/MM/yyyy");
  
// date du jour avec le format jour/mois/année pour la recherche du zip dans mon label hf-data-log-cmms
//  var date2 = date.getDate()-1;
  const folderId = 'ICI ID FOLDER;   //id du folder pour le dépot des pj
  const searchQuery = 'label:NOM-DU-LABEL '+realDate+' has:attachment';  //la recherche du ou des mails au préalable dans un label custom
const threads = GmailApp.search(searchQuery, 0, 10);
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const attachments = message.getAttachments({
          includeInlineImages: false,
          includeAttachments: true
      }); // la c'est la boucle avec les derniers params 
      attachments.forEach(attachment => {
        Drive.Files.insert(
          {
            title: attachment.getName(),
            mimeType: attachment.getContentType(),
            parents: [{ id: folderId }]
          },
          attachment.copyBlob()
        ); 
      });
    });
  });
};

function unzipFiles(){
  // Step 1. locate the zip file
  var ss = SpreadsheetApp.getActive();
  var zipType = "application/zip";
  
  var folderId = "ID DU DOSSIER";
  var zipFolder = DriveApp.getFolderById(folderId);
  
  var files = zipFolder.getFiles();
  
  while(files.hasNext()){
    var file = files.next();
     Logger.log(file);
    var fileType = file.getMimeType();
    if (fileType == zipType){
      // Step 2. Unzip the zip file
      var blob = file.getBlob().setContentTypeFromExtension();
      
      var unzippedFiles = Utilities.unzip(blob);
      
      for (var i in unzippedFiles){
        var unzippedFile = unzippedFiles[i];
        var fileName = unzippedFile.getName();
         Logger.log(fileName);
        fileName = fileName.split("/")[1];
        // Step 3. Convert the Excel file to Google Sheet
        var resource = {
          title: fileName,
          mimeType: MimeType.CSV
        };
        
        var mediaData = unzippedFile.copyBlob();
        var gs = Drive.Files.insert(resource, mediaData);
        var rootGS = DriveApp.getFileById(gs.id);
        // Step 4. Move the Google Sheet to destination folder and delete the original files from the root drive
        zipFolder.addFile(rootGS);
        DriveApp.removeFile(rootGS);
      }
      
    }
  }
  deleteZip();
}




function deleteZip(){
  var zipType = "application/zip";
  var folderId = "ID DOSSIER";
  var zipFolder = DriveApp.getFolderById(folderId);
  
  var files = zipFolder.getFiles();
  
  while(files.hasNext()){
    var file = files.next();
  
    var fileType = file.getMimeType();
    if (fileType == zipType){
       file.setTrashed(true);        // Dunk dans la corbeil les fichiers Zip du dossier !
    }  
  }
}





// extrac zip to drive <3 

//function unZip() {
//  //Add folder ID to select the folder where zipped files are placed
  var SourceFolder = DriveApp.getFolderById("ID DU DOSSIER");
//  //Add folder ID to save the where unzipped files to be placed
  var DestinationFolder = DriveApp.getFolderById("ID DU DOSSIER");
//  //Select the Zip files from source folder using the Mimetype of ZIP
  var ZIPFiles = SourceFolder.getFilesByType(MimeType.ZIP);
//  
//  //Loop over all the Zip files
  while (ZIPFiles.hasNext()){
//    
//   // Get the blob of all the zip files one by one
    var fileBlob = ZIPFiles.next().getBlob();
//   //Use the Utilities Class to unzip the blob
   var unZippedfile = Utilities.unzip(fileBlob);


   //Unzip the file and save it on destination folder
  var newDriveFile = DestinationFolder.createFile(unZippedfile[0]);
  }
 }
