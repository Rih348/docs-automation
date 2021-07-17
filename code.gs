const fields =  ['{{First Name}}', '{{Last Name}}','{{Email}}', '{{phone number}}', '{{Qualification}}', '{{Major}}', '{{Experience}}','{{Country}}' ];
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0]; // The first sheet
var sheet2 = ss.getSheets()[1]; // the En sheet
var sheet3 = ss.getSheets()[2]; // the Ar sheet

// This represents ALL the data
var range = sheet.getDataRange();
var values = range.getValues();

// This represents ALL the data in sheet2
var range_sheet2 = sheet2.getDataRange();
var values_sheet2 = range_sheet2.getValues();

// This represents ALL the data in sheet3
var range_sheet3 = sheet3.getDataRange();
var values_sheet3 = range_sheet3.getValues();


const target_folder = DriveApp.getFolderById('1_6gEo3FO2UpG0jPpvXQObTND-qym6Jmz');
const template_en_file = DriveApp.getFileById('1Ej-J2IS-NyB9psgCb-IQhJLHPtr4mC9yw52eISS1Qg8');
const template_ar_file = DriveApp.getFileById('1hlYvzvBNlfnRVP3fIAOx8T5N7UVoAvraX0vfi1ADQz8');

function CreateAllDocs() {

  // values[row][columnn]
  for (var i = 0 ; i < values.length; i= i+1){
    if (values[i][10] == '')
    {
      var filename = String(values[i][0]+' '+ values[i][1]);
      if (values[i][9] == 'Ar')
      {
      // // Creating AR Doc file 
      var document_copy = template_ar_file.makeCopy(filename);
      var document = DocumentApp.openById(document_copy.getId());

      bb = document.getBody();
      replacing(bb, i, 'Ar');

      }
      else
      {
      // // Creating EN Doc file  
      var document_copy = template_en_file.makeCopy(filename);
      var document = DocumentApp.openById(document_copy.getId());

      //Replace the text in the Doc
      bb = document.getBody();
      replacing(bb, i, 'En');

      }

      document_copy.moveTo(target_folder);

      // Updating the table 
      var value = 'created';
      var template_col_range = 'Sheet1!K'+ String(i+1);
      ss.getRange(template_col_range).setValue(value); // it return the range of the new data

    }

  }

}

function replacing(document, int,lang){
  if(lang == 'Ar')
  { 
    values = values_sheet3;
  }
  else 
  {
    values = values_sheet2; 
  }
  for (var j=0; j < 8; j= j+1)
  {
    document.replaceText(fields[j], values[int][j]);
  }
}

function onOpen() {
 SpreadsheetApp.getUi().createMenu("⚙️ Templates Creation")
   .addItem("Create All Docs", "CreateAllDocs")
   .addToUi();
}

