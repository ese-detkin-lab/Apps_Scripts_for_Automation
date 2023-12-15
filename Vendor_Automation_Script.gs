function extract_Vendor(e) {
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet = spreadSheet.getSheetByName('Form Responses 1');
  let digikey = spreadSheet.getSheetByName('Digikey');
  let mouser = spreadSheet.getSheetByName('Mouser');
  let adafruit = spreadSheet.getSheetByName('Adafruit');
  let mcm = spreadSheet.getSheetByName('McMaster-Carr');
  let amazon = spreadSheet.getSheetByName('Amazon');
  let sparkfun = spreadSheet.getSheetByName('Sparkfun');
  let misc = spreadSheet.getSheetByName('Misc');

  var columns = [5,11,17,23,29];
  var arr_cols = [10,11,12];

  var lastSourceRow = masterSheet.getLastRow();
  var lastSourceCol = masterSheet.getLastColumn()
  var rangemaster = masterSheet.getRange(1, 1, lastSourceRow, lastSourceCol);
  var data = rangemaster.getValues();

  //Logger.log(data); Debug for input data
  for(var i = 0;i<5;i++)
  {
  Logger.log('In Loop');
  //for (lastSourceRow-1 in data) {
    var targetValues = [];
    Logger.log(data[lastSourceRow-1][columns[i]]); //Debug for data being read
    if (data[lastSourceRow-1][columns[i]] == 'Amazon' || data[lastSourceRow-1][columns[i]] == 'amazon') {
      
      Logger.log('In Amazon');
      var tempvalue = [data[lastSourceRow-1][2],data[lastSourceRow-1][3],data[lastSourceRow-1][1],data[lastSourceRow-1][columns[i] + 1],data[lastSourceRow-1][columns[i] + 2],data[lastSourceRow-1][columns[i] + 3],data[lastSourceRow-1][columns[i] + 4],data[lastSourceRow-1][columns[i] + 5]];
      Logger.log(tempvalue);
      //then push that into the variables which holds all the new values to be returned
      targetValues.push(tempvalue);
      amazon.getRange(amazon.getLastRow() + 1, 2 , targetValues.length, 8).setValues(targetValues);

      ////// Code Snippet to add checkbox with each entry//////////////////
      var sheetlr = amazon.getLastRow();
      var cell1 = amazon.getRange(sheetlr,arr_cols[0]);
      cell1.insertCheckboxes();
      var cell2 = amazon.getRange(sheetlr,arr_cols[1]);
      cell2.insertCheckboxes();
      var cell3 = amazon.getRange(sheetlr,arr_cols[2]);
      cell3.insertCheckboxes();
      ///////////////////////////////////////////////////////////////////
    }
    else if (data[lastSourceRow-1][columns[i]] == 'Digikey' || data[lastSourceRow-1][columns[i]] == 'digikey' || data[lastSourceRow-1][columns[i]] == 'DigiKey') {
      Logger.log('In loop');
      var tempvalue = [data[lastSourceRow-1][2],data[lastSourceRow-1][3],data[lastSourceRow-1][1],data[lastSourceRow-1][columns[i] + 1],data[lastSourceRow-1][columns[i] + 2],data[lastSourceRow-1][columns[i] + 3],data[lastSourceRow-1][columns[i] + 4],data[lastSourceRow-1][columns[i] + 5]];
      Logger.log(tempvalue);
      //then push that into the variables which holds all the new values to be returned
      targetValues.push(tempvalue);
      digikey.getRange(digikey.getLastRow() + 1, 2 , targetValues.length, 8).setValues(targetValues);

      ////// Code Snippet to add checkbox with each entry//////////////////
      var sheetlr = digikey.getLastRow();
      var cell1 = digikey.getRange(sheetlr,arr_cols[0]);
      cell1.insertCheckboxes();
      var cell2 = digikey.getRange(sheetlr,arr_cols[1]);
      cell2.insertCheckboxes();
      var cell3 = digikey.getRange(sheetlr,arr_cols[2]);
      cell3.insertCheckboxes();
      ///////////////////////////////////////////////////////////////////

    }
    else if (data[lastSourceRow-1][columns[i]] == 'Mouser' || data[lastSourceRow-1][columns[i]] == 'mouser') {
      
      Logger.log('In loop');
      var tempvalue = [data[lastSourceRow-1][2],data[lastSourceRow-1][3],data[lastSourceRow-1][1],data[lastSourceRow-1][columns[i] + 1],data[lastSourceRow-1][columns[i] + 2],data[lastSourceRow-1][columns[i] + 3],data[lastSourceRow-1][columns[i] + 4],data[lastSourceRow-1][columns[i] + 5]];
      Logger.log(tempvalue);
      //then push that into the variables which holds all the new values to be returned
      targetValues.push(tempvalue);
      mouser.getRange(mouser.getLastRow() + 1, 2 , targetValues.length, 8).setValues(targetValues);

      ////// Code Snippet to add checkbox with each entry//////////////////
      var sheetlr = mouser.getLastRow();
      var cell1 = mouser.getRange(sheetlr,arr_cols[0]);
      cell1.insertCheckboxes();
      var cell2 = mouser.getRange(sheetlr,arr_cols[1]);
      cell2.insertCheckboxes();
      var cell3 = mouser.getRange(sheetlr,arr_cols[2]);
      cell3.insertCheckboxes();
      ///////////////////////////////////////////////////////////////////
    }
    else if (data[lastSourceRow-1][columns[i]] == 'Adafruit' || data[lastSourceRow-1][columns[i]] == 'adafruit') {
      
      Logger.log('In loop');
      var tempvalue = [data[lastSourceRow-1][2],data[lastSourceRow-1][3],data[lastSourceRow-1][1],data[lastSourceRow-1][columns[i] + 1],data[lastSourceRow-1][columns[i] + 2],data[lastSourceRow-1][columns[i] + 3],data[lastSourceRow-1][columns[i] + 4],data[lastSourceRow-1][columns[i] + 5]];
      Logger.log(tempvalue);
      //then push that into the variables which holds all the new values to be returned
      targetValues.push(tempvalue);
      adafruit.getRange(adafruit.getLastRow() + 1, 2 , targetValues.length, 8).setValues(targetValues);

      ////// Code Snippet to add checkbox with each entry//////////////////
      var sheetlr = adafruit.getLastRow();
      var cell1 = adafruit.getRange(sheetlr,arr_cols[0]);
      cell1.insertCheckboxes();
      var cell2 = adafruit.getRange(sheetlr,arr_cols[1]);
      cell2.insertCheckboxes();
      var cell3 = adafruit.getRange(sheetlr,arr_cols[2]);
      cell3.insertCheckboxes();
      ///////////////////////////////////////////////////////////////////
    }
    else if (data[lastSourceRow-1][columns[i]] == 'Mcmaster-Carr' || data[lastSourceRow-1][columns[i]] == 'MCM') {
      //Save it ta a temporary variable
      Logger.log('In loop');
      var tempvalue = [data[lastSourceRow-1][2],data[lastSourceRow-1][3],data[lastSourceRow-1][1],data[lastSourceRow-1][columns[i] + 1],data[lastSourceRow-1][columns[i] + 2],data[lastSourceRow-1][columns[i] + 3],data[lastSourceRow-1][columns[i] + 4],data[lastSourceRow-1][columns[i] + 5]];
      Logger.log(tempvalue);
      //then push that into the variables which holds all the new values to be returned
      targetValues.push(tempvalue);
      mcm.getRange(mcm.getLastRow() + 1, 2 , targetValues.length, 8).setValues(targetValues);

      ////// Code Snippet to add checkbox with each entry//////////////////
      var sheetlr = mcm.getLastRow();
      var cell1 = mcm.getRange(sheetlr,arr_cols[0]);
      cell1.insertCheckboxes();
      var cell2 = mcm.getRange(sheetlr,arr_cols[1]);
      cell2.insertCheckboxes();
      var cell3 = mcm.getRange(sheetlr,arr_cols[2]);
      cell3.insertCheckboxes();
      ///////////////////////////////////////////////////////////////////
    }
    else if (data[lastSourceRow-1][columns[i]] == 'Sparkfun' || data[lastSourceRow-1][columns[i]] == 'sparkfun') {
      //Save it ta a temporary variable
      Logger.log('In loop');
      var tempvalue = [data[lastSourceRow-1][2],data[lastSourceRow-1][3],data[lastSourceRow-1][1],data[lastSourceRow-1][columns[i] + 1],data[lastSourceRow-1][columns[i] + 2],data[lastSourceRow-1][columns[i] + 3],data[lastSourceRow-1][columns[i] + 4],data[lastSourceRow-1][columns[i] + 5]];
      Logger.log(tempvalue);
      //then push that into the variables which holds all the new values to be returned
      targetValues.push(tempvalue);
      sparkfun.getRange(sparkfun.getLastRow() + 1, 2 , targetValues.length, 8).setValues(targetValues);

      ////// Code Snippet to add checkbox with each entry//////////////////
      var sheetlr = sparkfun.getLastRow();
      var cell1 = sparkfun.getRange(sheetlr,arr_cols[0]);
      cell1.insertCheckboxes();
      var cell2 = sparkfun.getRange(sheetlr,arr_cols[1]);
      cell2.insertCheckboxes();
      var cell3 = sparkfun.getRange(sheetlr,arr_cols[2]);
      cell3.insertCheckboxes();
      ///////////////////////////////////////////////////////////////////
    } 
    else if(data[lastSourceRow-1][columns[i]] == 'Miscellaneous'){
      //Save it ta a temporary variable
      Logger.log('In loop');
      var tempvalue = [data[lastSourceRow-1][2],data[lastSourceRow-1][3],data[lastSourceRow-1][1],data[lastSourceRow-1][columns[i] + 1],data[lastSourceRow-1][columns[i] + 2],data[lastSourceRow-1][columns[i] + 3],data[lastSourceRow-1][columns[i] + 4],data[lastSourceRow-1][columns[i] + 5]];
      Logger.log(tempvalue);
      //then push that into the variables which holds all the new values to be returned
      targetValues.push(tempvalue);
      misc.getRange(misc.getLastRow() + 1, 2 , targetValues.length, 8).setValues(targetValues);

      ////// Code Snippet to add checkbox with each entry//////////////////
      var sheetlr = misc.getLastRow();
      var cell1 = misc.getRange(sheetlr,arr_cols[0]);
      cell1.insertCheckboxes();
      var cell2 = misc.getRange(sheetlr,arr_cols[1]);
      cell2.insertCheckboxes();
      var cell3 = misc.getRange(sheetlr,arr_cols[2]);
      cell3.insertCheckboxes();
      ///////////////////////////////////////////////////////////////////
    }
  }
}

// Add checkbox example function , keep this commented. For reference only
/*function add_checkbox(e){
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet = spreadSheet.getSheetByName('Form Responses 1');

  var arr_cols = [12,19,26,33,40];

  var lastSourceRow = masterSheet.getLastRow();

  for(var i = 0;i<5;i++)
  { 
    var cell = masterSheet.getRange(lastSourceRow,arr_cols[i]);
    cell.insertCheckboxes();
  }

}*/

function sendEmail(e){

    var sheet = e.source.getActiveSheet();
    var cell = e.range;
    var lastSourceRow = cell.getRow();

    var arr1 = sheet.getRange(lastSourceRow,11);

  if(arr1.isChecked())
  {
      var reciever = sheet.getRange(lastSourceRow,4).getValues();
      var part = sheet.getRange(lastSourceRow,6).getValues();

      Logger.log(reciever);
      MailApp.sendEmail({
        to: reciever.toString(),
        subject: "Item with Catalog Number: " + part + " has arrived",
        body: "Your Part with Catalog Number: " + part + " is here. Pick it up from Detkin.",
      })
  }
}












