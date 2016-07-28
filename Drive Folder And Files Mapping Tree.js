
/* 
This code with spreadheet create Chart Tree of Folders directory and All Files with URL link and Custome menu to
revoke access and permissions for files and Folders. 
Here is dummy spreadsheet https://docs.google.com/spreadsheets/d/1tF8RYThzLvgcOVxGBfR0n2csQwPiU50k46VD5qwMWyY/edit?usp=sharing
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Run Macro')
  .addItem('Get All Sub Folders','getFoldersandParent')
  .addItem('Get All Files', 'getAllFilesAndPath')
  .addSeparator()
  .addSubMenu(ui.createMenu('Remove Access')
              .addItem('Make All Folder Private', 'revokFolderAccess')
              .addItem('Make All Files Private', 'revokFileAccess'))
  .addToUi();
}


function getFoldersandParent(){
  removeDuplicates("Folder")
  var ss = SpreadsheetApp.getActive(), id;
  var sheet = ss.getSheetByName("Folder"), data; sheet.getRange(2, 1).setFormula("=Directory!B1");
  sheet.getRange(2, 2).setValue("My Drive"),sheet.getRange(2, 7).setFormula("=COUNTA(E:E)+1")
  sheet.getRange(2, 12).setFormula("=COUNTA(K:K)+1")
  sheet.getRange(2, 8).setFormula("=COUNTA(F:F)+1"),sheet.getRange(2, 10).setFormula("=COUNTA(I:I)+1");
  var i = sheet.getRange(2, 7).getValue();
  var parent = DriveApp.getFoldersByName(sheet.getRange(2, 1).getValue()).next();
  sheet.getRange(2, 4).setValue(parent.getId());sheet.getRange(2, 3).setValue(parent.getUrl())
  while(sheet.getRange(i, 4).getValue() != ""){
    id = sheet.getRange(i, 4).getValue()
    var folder = DriveApp.getFolderById(id).getFolders()
    while(folder.hasNext()){
      var nestFoder = folder.next()
      data=[
        nestFoder.getName(),
        nestFoder.getParents().next(),
        nestFoder.getUrl(),
        nestFoder.getId()];
      sheet.appendRow(data)
    }
    sheet.getRange(i, 5).setValue("G");
    i++;
  }
  validateFolder();
  getAllFilesAndPath();
  validateFiles();
}


function getAllFilesAndPath(){
  removeDuplicates("Directory")
  var ss = SpreadsheetApp.getActive();
  var sheet1 = ss.getSheetByName("Folder"),sheet2 = ss.getSheetByName("Directory");
  sheet2.getRange(2, 6).setFormula("=COUNTA(E:E)+1")
  var i = sheet1.getRange(2, 8).getValue(), data;
  while(sheet1.getRange(i, 4).getValue() != "" ){
    var folderID = sheet1.getRange(i, 4).getValue();
    var folder = DriveApp.getFolderById(folderID);
    var files = folder.getFiles()
    while(files.hasNext()){
      var file = files.next();
      data = [
        file.getName(),
        file.getUrl(),
        getFullFilesrPath(file.getId()),
        file.getId()
      ];
        sheet2.appendRow(data);
    }
    sheet1.getRange(i, 6).setValue("C");
    i++;
  }
  
}

function getFullFilesrPath(id) {
  // For identifieng parent folder
  var path, folder = DriveApp.getFileById(id);
  var folders = [],
      parent = folder.getParents();
  while (parent.hasNext()) {
    parent = parent.next();
      folders.push(parent);
      parent = parent.getParents();
  }
  
  if (folders.length) {
    folders = folders.reverse()
    path = folders//.reverse()
    var filePath  = folders.join("/")
    return (filePath)
    }
}

function removeDuplicates(name) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(name);
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

function revokFolderAccess(){
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Folder");
  var i = sheet.getRange(2, 10).getValue();
  while(sheet.getRange(i, 4).getValue() != ""){
    var folder = DriveApp.getFolderById(sheet.getRange(i, 4).getValue())
    folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    var vw = folder.getViewers();
    var ed = folder.getEditors();
    var ln = vw.length;
    for (var j=0;j<ln;j++)
    {
      var Uid = vw[j].getEmail();
      folder.removeViewer(Uid);
    }
    var ln = ed.length;
    for (var k=0;k<ln;k++)
    {
      var Uid = ed[k].getEmail();
      if (Uid !="" )
      folder.removeEditor(Uid);
    }
    sheet.getRange(i, 9).setValue("P")
    i++;
  }

}

function revokFileAccess(){
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Directory");
  sheet.getRange(2, 6).setFormula("=COUNTA(E:E)+2")
  var i = sheet.getRange(2, 6).getValue()
  while(sheet.getRange(i, 4).getValue() != ""){
    var ID = sheet.getRange(i, 4).getValue();
    var file = DriveApp.getFileById(ID);
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    var vw = file.getViewers();
    var ed = file.getEditors();

    var ln = vw.length;
    for (var j=0;j<ln;j++)
    {
      var Uid = vw[j].getEmail();
      file.removeViewer(Uid);
    }
  
    var ln = ed.length;
    for (var k=0;k<ln;k++)
    {
      var Uid = ed[k].getEmail();
      Logger.log(Uid)
      if (Uid !="" )
      file.removeEditor(Uid);
    }
    sheet.getRange(i, 5).setValue("P");
    i++;
  }
}

function validateFolder(){
  var ss = SpreadsheetApp.getActive(), i = 3;
  var sheet = ss.getSheetByName("Folder"), realParent = sheet.getRange(2, 1).getValue();
  sheet.getRange(2, 12).setFormula("=COUNTA(K:K)+1")
  var i = sheet.getRange(2, 12).getValue()
  sheet.getRange(2, 11).setValue("V")
  while(sheet.getRange(i, 4).getValue() != ""){
    var folder = DriveApp.getFolderById(sheet.getRange(i, 4).getValue())
    var folders = [],
        parent = folder.getParents();
    while (parent.hasNext()) {
      parent = parent.next();
      folders.push(parent);
      parent = parent.getParents();
    }
    folders.reverse();
    folders.push(sheet.getRange(i, 1).getValue());
    Logger.log(folders)
    if(folders[1] != realParent){
      sheet.deleteRow(i); i--;
    }
    else{sheet.getRange(i, 11).setValue("V")}
    i++;
  }
}


function validateFiles(){
  var ss = SpreadsheetApp.getActive(), i = 3;
  var sheet = ss.getSheetByName("Directory"), realParent = sheet.getRange(1, 2).getValue();
  sheet.getRange(2, 12).setFormula("=COUNTA(G:G)+2")
  var i = sheet.getRange(2, 12).getValue()
  while(sheet.getRange(i, 4).getValue() != ""){
    var folder = DriveApp.getFileById(sheet.getRange(i, 4).getValue())
    var folders = [],
        parent = folder.getParents();
    while (parent.hasNext()) {
      parent = parent.next();
      folders.push(parent);
      parent = parent.getParents();
    }
    folders.reverse();
    Logger.log(folders[1])
    if(folders[1] != realParent){
      sheet.deleteRow(i); i--;
    }
    else{sheet.getRange(i, 7).setValue("V")}
    i++;
  }
}
