//always call SpreadsheetApp.flush() before passing the sheet to this function.
//pass a google spreadsheet object to this function, an array of emails [single@email] is acceptable as long as its an array. set trash to true to delete original, false to not touch.
//reportname becomes the subject
function emailAsXlsx(ss,emails,reportname,trashed){
  url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + ss.getId() + "&exportFormat=xlsx";
  var params = {
      method      : "get",
      headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    };
  var blob = UrlFetchApp.fetch(url, params).getBlob();  
  blob.setName(ss.getName() + '.xlsx');
  for(var i = 0; i < emails.length; i++){
    MailApp.sendEmail(emails[i],reportname,"Please see attached report, this is an automated email", {attachments: [blob]});
  }
  if(trashed){
    DriveApp.getFileById(ss.getId()).setTrashed(true);
  }
} 


//this function will convert a spreadsheet into a pdf, save it in the folder id passed to it, name it per the name variable, and trash the ss if you set trashed to true. 
//returns the pdf as a blob for further processing.
function exportPdf(ss,folder,name,trashed){
  var url = ss.getUrl();
  url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/"
  var url_ext = 'export?exportFormat=pdf&format=pdf' + 
    '&size=letter' + 
    '&portrait=false' + 
    '&fitw=true' + 
    '&scale=4' +   
    '&sheetnames=false&printtitle=false&pagenumbers=false' + 
    '&gridlines=false' + 
    '&fzr=false' + 
    '&id=' + ss.getId(); 
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url + url_ext, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  var blob = response.getBlob().setName(name + '.pdf');
  var tfolder = DriveApp.getFolderById(folder);
  var file = tfolder.createFile(blob);
  if(trashed){
    DriveApp.getFileById(ss.getId()).setTrashed(true);
  }
  return file.getBlob();
}  


//this is an example of how to get the a1 notation for an array for bigWrite
//var range = dsheet.getRange(1,1,data.length,data[0].length).getA1Notation();

//this function, as the name implies, writes huge chunks of data. It is much more efficient than the getRange.setValues() option for huge arrays. 
//requires sheets api enabled for project.
//spreadsheet ID 'spreadsheet object'.getId() var ssid = SpreadsheetApp.getActiveSpreadsheet().getId()
//data array to be written
//range A1 notation of the range you want to write the array to, this MUST match the array size or it will fail, its this way so you dont have to write to a clean sheet, and can update part of it
//sheet name string representing the destination sheet, destination sheet must be on the spreadsheet represented by the spreadsheet id
function bigWrite(data,ssid,sheetname,range){
  var sheet = SpreadsheetApp.openById(ssid).getSheetByName(sheetname);
  sheet.getRange(range).clearContent();
  SpreadsheetApp.flush();
  var fullrange = sheetname+'!'+range;
  var request = {
  'range': fullrange,
  'majorDimension': 'ROWS',
  'values': data
  }
  Sheets.Spreadsheets.Values.update(request,ssid,fullrange,{'valueInputOption': 'RAW'});
}