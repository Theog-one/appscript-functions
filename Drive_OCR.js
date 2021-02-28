function dumpPdfText(fileId){
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var resource = {title: blob.getName(),mimeType: blob.getContentType()};
  var ocrfile = Drive.Files.insert(resource, blob, {ocr: true, ocrLanguage: "en"});
  var document = DocumentApp.openById(ocrfile.id);
  var text = document.getBody().getText();
  DriveApp.getFileById(document.getId()).setTrashed(true);
  return text;
}

function sortPdf(){
  var wodata = SpreadsheetApp.openById('').getSheetByName('').getDataRange().getDisplayValues();
  var folder = DriveApp.getFolderById('');
  var destFolder = DriveApp.getFolderById('');
  var files = folder.getFiles();
  while(files.hasNext()){
    var file = files.next();
    var text = dumpPdfText(file.getId());
    var name = matchWoNum(wodata,text);
    file.setName(name);
    if(name != '??????'){
      file.moveTo(destFolder);
    }
  }
}

function matchWoNum(wodata,text){
  var wonums = [];
  for(var i = 0; i < wodata.length; i++){
    var wonum = wodata[i][0];
    if(wonum.length > 6){
      wonum = wonum.slice(0,-1);
    }
    if(text.search(wonum) != -1){
      wonums.push(wonum);
    }
  }
  if(wonums.length > 0){
    if(wonums.length > 1){
      return wonums[0] + '?';
    }
    else{
      return wonums[0];
    }
  }
  else{
    return '??????'
  }
}

