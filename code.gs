//function to move screen where user can select available options
function jumpToCell() { 
  var s = SpreadsheetApp.getActiveSpreadsheet(); 
  var tab = s.getSheetByName("Account Summary"); 
  tab.setActiveCell(tab.getRange(59, 1)); 
}
 
//Copy and paste with width and height Values: function to copy original template to TEMP sheet
function copyAllInRange(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var defSheet = ss.getSheetByName('Account Summary');
var tempSheet = ss.getSheetByName('TEMP');
// Get the range from the sheet you will copy from. 
var defaultRange = defSheet.getRange(1, 1, 68, 10);
// Get the range of the sheet that is your paste target.
var pasteRange = tempSheet.getRange(1, 1, 68, 10);
var columnWidths = SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS
defaultRange.copyTo(pasteRange);
defaultRange.copyTo(pasteRange,columnWidths,false);  
};

//function to reset the script to the original state 
function resetAllRange(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var defSheet = ss.getSheetByName('Account Summary');
var tempSheet = ss.getSheetByName('TEMP');
var pdfSheet = ss.getSheetByName('ScriptSheet');
pdfSheet.clear();
// Get the range from the sheet you will copy from. 
var defaultRange = tempSheet.getRange(1, 1, 68, 10);
// Get the range of the sheet that is your paste target.
var pasteRange = defSheet.getRange(1, 1, 68, 10);
var columnWidths = SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS
defaultRange.copyTo(pasteRange);
defaultRange.copyTo(pasteRange,columnWidths,false);
var lastRow = defSheet.getLastRow();
defSheet.getRange(68,1,lastRow,10).clear();
allFiles = DriveApp.getFilesByName("_AS Imported.xls");
while (allFiles.hasNext()) {//If there is another element in the iterator
thisFile = allFiles.next();
idToDLET = thisFile.getId();  
Drive.Files.remove(idToDLET);
};
allFiles1 = DriveApp.getFilesByName("_AS Imported.xls.csv");
while (allFiles1.hasNext()) {//If there is another element in the iterator
thisFile1 = allFiles1.next();
idToDLET1 = thisFile1.getId();  
Drive.Files.remove(idToDLET1);
};
};

//function to compress the invoice and remove empty rows
function deleteCells() { 
var ss = SpreadsheetApp.getActiveSpreadsheet();  
var s = ss.getSheetByName('Account Summary');
var r = s.getRange('J:J');    
var v = r.getValues(); 
//SpreadsheetApp.getUi().alert(v); 
for(var i=v.length;i>0;i--) 
{
if(v[0,i]=='false')
{ 
s.deleteRow(i+1);
}
}
var values = s.getRange('A1:A70').getValues();
for(var i=values.length;i>12;i--)
{ 
if(values[0,i]=='')
{
s.deleteRow(i+1);
}
}
};  

//function that copies the invoice to another sheet and saves pdf and xls copy to the computer
function downloadSheetAsPDF2() {  
var sheet = SpreadsheetApp.getActiveSpreadsheet();
var defSheet = sheet.getSheetByName('Account Summary');
var tempSheet = sheet.getSheetByName('ScriptSheet');
var lastRow = sheet.getLastRow();
// Get the range from the sheet you will copy from. 
var defaultRange = defSheet.getRange(1, 1, lastRow, 9);

var row_heights = getRowHeights(defaultRange);

// Get the range of the sheet that is your paste target.
var pasteRange = tempSheet.getRange(1, 1, lastRow, 9);
var columnWidths = SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS
defaultRange.copyTo(pasteRange);
defaultRange.copyTo(pasteRange,columnWidths,false);
 var set_row_heights = setRowHeights(row_heights,pasteRange);

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetId = ss.getSheetByName('ScriptSheet').getSheetId();
var filename = ss.getRange('B7').getValue();
// Creat XLS file as a temporary file and create URL for downloading.
var s1= ss.getSheetByName('ScriptSheet');
s1.setHiddenGridlines(false);

var url1 = "https://docs.google.com/a/mydomain.org/spreadsheets/d/" + ss.getId() + "/export?exportFormat=xlsx&gid=" + sheetId + "&access_token=" + ScriptApp.getOAuthToken();
var blob1 = UrlFetchApp.fetch(url1).getBlob().setName(filename + '- Account Summary.xlsx');
var file1 = DriveApp.createFile(blob1);
var dlUrl1 = "https://drive.google.com/uc?export=download&id=" + file1.getId();

var str1 = '<script>window.location.href="' + dlUrl1 + '"</script>';
var html1 = HtmlService.createHtmlOutput(str1);
SpreadsheetApp.getUi().showModalDialog(html1, "Downloading...");
file1.setTrashed(true);
  
// This is used for closing the dialog.
Utilities.sleep(3000);
var closeHtml = HtmlService.createHtmlOutput("<script>google.script.host.close()</script>");
SpreadsheetApp.getUi().showModalDialog(closeHtml, "Downloading...");

// Creat PDF file as a temporary file and create URL for downloading.
s1.setHiddenGridlines(true);
var url = "https://docs.google.com/a/mydomain.org/spreadsheets/d/" + ss.getId() + "/export?exportFormat=pdf&gid=" + sheetId + "&access_token=" + ScriptApp.getOAuthToken();
var blob = UrlFetchApp.fetch(url).getBlob().setName(filename + '- Account Summary.pdf');
var file = DriveApp.createFile(blob);
var dlUrl = "https://drive.google.com/uc?export=download&id=" + file.getId();

// Open a dialog and run Javascript for downloading the file.
var str = '<script>window.location.href="' + dlUrl + '"</script>';
var html = HtmlService.createHtmlOutput(str);
SpreadsheetApp.getUi().showModalDialog(html, "Downloading...");
file.setTrashed(true);

// This is used for closing the dialog.
Utilities.sleep(3000);
var closeHtml = HtmlService.createHtmlOutput("<script>google.script.host.close()</script>");
SpreadsheetApp.getUi().showModalDialog(closeHtml, "Downloading...");
}

//function that converts the xls file to csv, loads it into template and checks values
function convertExceltoGoogleSpreadsheet2(fileName) {
  try {
    fileName = fileName || "_AS Imported.xls";
    var oauthToken = ScriptApp.getOAuthToken();
    var excelFile = DriveApp.getFilesByName(fileName).next();
    var fileId = excelFile.getId();
    var folderId = Drive.Files.get(fileId).parents[0].id;
    var targetFolder = DriveApp.getFolderById(folderId);
    // Re-upload the XLS file after convert in Google Sheet format
            var googleSheet = JSON.parse(UrlFetchApp.fetch(
                "https://www.googleapis.com/upload/drive/v2/files?uploadType=media&convert=true", {
                    method: "POST",
                    contentType: "application/vnd.ms-excel",
                    payload: excelFile.getBlob().getBytes(),
                    headers: {
                        "Authorization": "Bearer " + oauthToken
                    }
                }
            ).getContentText());

    // The exportLinks object has a link to the converted CSV file
            var targetFile = UrlFetchApp.fetch(
                googleSheet.exportLinks["text/csv"], {
                    method: "GET",
                    headers: {
                        "Authorization": "Bearer " + oauthToken
                    }
                });
                
    // Save the CSV file in the destination folder
            targetFolder.createFile(targetFile.getBlob()).setName(excelFile.getName() + ".csv");

  } catch (f) {
    Logger.log(f.toString());
  }

 var file = DriveApp.getFilesByName("_AS Imported.xls.csv").next();
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(100, 1, csvData.length, csvData[0].length).setValues(csvData);

sheet.getRange(5,2).setValue(sheet.getRange(101,2).getValue());//date
sheet.getRange(7,2).setValue(sheet.getRange(102,2).getValue());//name
sheet.getRange(8,2).setValue(sheet.getRange(103,2).getValue());//dol
sheet.getRange(9,2).setValue(sheet.getRange(104,2).getValue());//claim #
sheet.getRange(7,8).setValue(sheet.getRange(102,8).getValue());//fdot
sheet.getRange(8,8).setValue(sheet.getRange(103,8).getValue());//ldot

var lastRow = sheet.getLastRow();
var numRows = lastRow-105;
var content = sheet.getRange(106,1,numRows,8).getValues();
sheet.getRange(12, 1, numRows, 8).setValues(content);
// Money format
var range = sheet.getRange(12, 6, 32, 9);
range.setNumberFormat("$#,##0.00;$(#,##0.00)");
sheet.getRange(68,1,lastRow,10).clear();

var lastRow = sheet.getLastRow();
var ocfs = sheet.getRange('A12:A31');
var vs = ocfs.getValues();
for (var n = vs.length-1; n>=0; n--)
{

if(vs[0,n].toString().indexOf("OCF3") > -1)
{
//SpreadsheetApp.getUi().alert(vs[0,n]); 
sheet.getRange(n+12,6).setValue(200); 
sheet.getRange(n+12,7).setValue(200);
sheet.getRange(n+12,8).setBackground("red"); 
}
if(vs[0,n].toString().indexOf("CNR") > -1)
{
sheet.getRange(n+12,5).setValue('Sent'); 

}
if(vs[0,n].toString().indexOf("OCF23") > -1)
{ 
sheet.getRange(n+12,7).setBackground("red");
sheet.getRange(n+12,8).setBackground("red"); 
}
if(vs[0,n] == 'OCF18 - 1')
{ 
if (sheet.getRange(n+12,5).getValue() == 'Not Approved')
{
sheet.getRange(n+12,7).setBackground("red");
}
}

var partial = sheet.getRange('E12:E31');
var vals = partial.getValues();
for (var k = vals.length-1; k>=0; k--)
{
if(vals[0,k].toString().indexOf("Partially Approved") > -1)
{ 
sheet.getRange(k+12,5).setValue("Partially Approved for $"); 
sheet.getRange(k+12,5).setWrap(true);
sheet.getRange(k+12,5).setBackground("red"); 
}
}
}
}


/**
 * GET THE ROW HEIGHTS OF A SELECTED RANGE OF A SOURCE SHEET
 * @param {object} range - Selected source range
 * @returns {Array.<number>}  Array of row heights for each row
 */
function getRowHeights(range) {
 
 var rngRowStart = range.getRow();
 var rngRowHeight = range.getHeight() + rngRowStart;
 
 var rowHeights = []
 var rangeSheet = range.getSheet()
 
 for( var i = rngRowStart; i < rngRowHeight; i++){
   var rowHeight = rangeSheet.getRowHeight(i);
   rowHeights.push(rowHeight);
 };
 return rowHeights;
};
 
 
/**
 * SET THE ROW HEIGHTS OF A SELECTED RANGE OF A DESTINATION SHEET
 * @param {Array.<number>} rHeights - row heights from getRowHeights(rng);
 * @param {object} destRange - destionation range of copied data.
 */
function setRowHeights(rHeights,destRange){
 
 var rngRowStart = destRange.getRow();
 var rngRowHeight = destRange.getHeight() + rngRowStart;
 
 var destSheet = destRange.getSheet();
 Logger.log(destSheet.getName());
 var count = 0;
 for( var i = rngRowStart; i < rngRowHeight; i++){
   destSheet.setRowHeight(i,rHeights[count]);
   
   count+=1
 };
};


//function that changes highlighted cell back on edit
function onEdit(e){
var ss = SpreadsheetApp.getActiveSpreadsheet();  
var s = ss.getSheetByName('Account Summary');
var range = e.range;
var color = range.getBackground();
var col = range.getColumn();
var row = range.getRow(); 
var adj = s.getRange(row, col-1);
var adjColor = adj.getBackground();
//SpreadsheetApp.getUi().alert(color);
  if(color =='#ff0000'){
    range.setBackground(adjColor);
  }
}

//functions to upload the source data from the computer
function doPost(e) {
var data = Utilities.base64Decode(e.parameters.data);
var blob = Utilities.newBlob(data, e.parameters.mimetype, e.parameters.filename);
DriveApp.createFile(blob);
var output = HtmlService.createHtmlOutput("<b>Done!</b>");
output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
return output;
}
function openDialog() {
var html = HtmlService.createHtmlOutputFromFile('Index');
SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Upload _AS Imported.xls');
}
function doGet() {
return HtmlService.createHtmlOutputFromFile('Index');
}
