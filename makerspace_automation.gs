function service_sort(e) {
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet = spreadSheet.getSheetByName('Form Responses 1');
  let printing = spreadSheet.getSheetByName('3D-Printing');
  let cutting = spreadSheet.getSheetByName('Laser-Cutting');

  var serv_req = 5
  var error_col_print = [10,11,12,13,14]
  var error_col_cutting = [9,10,11,12,13]

  var lastSourceRow = masterSheet.getLastRow();
  var lastSourceCol = masterSheet.getLastColumn()
  var rangemaster = masterSheet.getRange(1, 1, lastSourceRow, lastSourceCol);
  var data = rangemaster.getValues();

  var targetValues = [];

  if (data[lastSourceRow-1][serv_req] == '3D Printing') 
  {
    var tempvalue = [data[lastSourceRow-1][1],data[lastSourceRow-1][2],data[lastSourceRow-1][4],data[lastSourceRow-1][6],data[lastSourceRow-1][7],data[lastSourceRow-1][8],data[lastSourceRow-1][9],data[lastSourceRow-1][10]];

    // Pushing datastructure in new sheet
    targetValues.push(tempvalue);
    printing.getRange(printing.getLastRow() + 1, 2 , targetValues.length, 8).setValues(targetValues);

    // Adding the Error Reporting Check Boxes
    var sheetlr = printing.getLastRow();
    //Printing ?
    var cell1 = printing.getRange(sheetlr,error_col_print[0]);
    cell1.insertCheckboxes();
    //Finished ?
    var cell2 = printing.getRange(sheetlr,error_col_print[1]);
    cell2.insertCheckboxes();
    //incompatible dimensions ?
    var cell3 = printing.getRange(sheetlr,error_col_print[2]);
    cell3.insertCheckboxes();
    //wrong file type ?
    var cell4 = printing.getRange(sheetlr,error_col_print[3]);
    cell4.insertCheckboxes();
    //Queue Full ?
    var cell5 = printing.getRange(sheetlr,error_col_print[4]);
    cell5.insertCheckboxes();
  }
  else if(data[lastSourceRow-1][serv_req] == 'Laser Cutting')
  {
    var tempvalue = [data[lastSourceRow-1][1],data[lastSourceRow-1][2],data[lastSourceRow-1][4],data[lastSourceRow-1][11],data[lastSourceRow-1][12],data[lastSourceRow-1][13],data[lastSourceRow-1][14]];

    // Pushing datastructure in new sheet
    targetValues.push(tempvalue);
    cutting.getRange(cutting.getLastRow() + 1, 2 , targetValues.length, 7).setValues(targetValues);

    // Adding the Error Reporting Check Boxes
    var sheetlr = cutting.getLastRow();
    //Printing ?
    var cell1 = cutting.getRange(sheetlr,error_col_cutting[0]);
    cell1.insertCheckboxes();
    //Finished ?
    var cell2 = cutting.getRange(sheetlr,error_col_cutting[1]);
    cell2.insertCheckboxes();
    //incompatible dimensions ?
    var cell3 = cutting.getRange(sheetlr,error_col_cutting[2]);
    cell3.insertCheckboxes();
    //wrong file type ?
    var cell4 = cutting.getRange(sheetlr,error_col_cutting[3]);
    cell4.insertCheckboxes();
    //Queue Full ?
    var cell5 = cutting.getRange(sheetlr,error_col_cutting[4]);
    cell5.insertCheckboxes();
    
  }
    
}

function sendEmail(e){

    var sheet = e.source.getActiveSheet();
    var cell = e.range;
    var lastSourceRow = cell.getRow();

  if(sheet.getName() == '3D-Printing')   
  {

    var arr1 = sheet.getRange(lastSourceRow,11);
    var arr2 = sheet.getRange(lastSourceRow,12);
    var arr3 = sheet.getRange(lastSourceRow,13);
    var arr4 = sheet.getRange(lastSourceRow,14);
    var reciever = sheet.getRange(lastSourceRow,2).getValues();
    var course = sheet.getRange(lastSourceRow,3).getValues();
    if(arr1.isChecked())
    {

      //Logger.log(reciever);
      MailApp.sendEmail({
        to: reciever.toString(),
        subject: "3D Printing Request Complete",
        body: 'Hello,3D priting request has successfully completed for course: ' + course + '. Please collect it from Detkin.',
      })
    }
    else if(arr2.isChecked())
    {
          //Logger.log(reciever);
      MailApp.sendEmail({
        to: reciever.toString(),
        subject: "Service Request Denied",
        body: 'Hello,Your 3D Printing Request for course: ' + course + ' is Denied. The Dimensions submitted are too big for our printers.Please explore alternative options like Tangen, RPL etc.' ,
      })
    }
    else if(arr3.isChecked())
    {
          //Logger.log(reciever);
      MailApp.sendEmail({
        to: reciever.toString(),
        subject: "Service Request Denied",
        body: "Hello,Your 3D Printing Request for course:" + course + " is Denied. The file submitted is of wrong type. Please resubmit the form with correct file type." ,
      })
    }
    else if(arr4.isChecked())
    {
          //Logger.log(reciever);
      MailApp.sendEmail({
        to: reciever.toString(),
        subject: "Service Request Denied",
        body: "Hello,Your 3D Printing Request for course:" + course + " is Denied. Detkin printers are currently busy with previous requests. Please explore alternative options like Tangen, RPL etc." ,
      })
    }
  }
  else if(sheet.getName() == 'Laser-Cutting')
  {

    var arr1 = sheet.getRange(lastSourceRow,10);
    var arr2 = sheet.getRange(lastSourceRow,11);
    var arr3 = sheet.getRange(lastSourceRow,12);
    var arr4 = sheet.getRange(lastSourceRow,13);
    var reciever = sheet.getRange(lastSourceRow,2).getValues();
    var course = sheet.getRange(lastSourceRow,3).getValues();

    if(arr1.isChecked())
    {

      //Logger.log(reciever);
      MailApp.sendEmail({
        to: reciever.toString(),
        subject: "Laser Cutting Request Complete",
        body: 'Hello,Laser Cutting request has successfully completed for course: ' + course + '. Please collect it from Detkin.' ,
      })
    }
    else if(arr2.isChecked())
    {
          //Logger.log(reciever);
      MailApp.sendEmail({
        to: reciever.toString(),
        subject: "Service Request Denied",
        body: 'Hello,Your Laser Cutting Request for course:' + course + 'is Denied. The Dimensions submitted are too big for our Laser Cutter.Please explore alternative options like Tangen, RPL etc.' ,
      })
    }    
    else if(arr3.isChecked())
    {
          //Logger.log(reciever);
      MailApp.sendEmail({
        to: reciever.toString(),
        subject: "Service Request Denied",
        body: "Hello,Your Laser Cutting Request for course:" + course + " is Denied. The file submitted is of wrong type. Please resubmit the form with correct file type." ,
      })
    }
    else if(arr4.isChecked())
    {
          //Logger.log(reciever);
      MailApp.sendEmail({
        to: reciever.toString(),
        subject: "Service Request Denied",
        body: "Hello,Your Laser Cutting Request for course:" + course + " is Denied. Detkin Laser Cutter is currently busy with previous requests. Please explore alternative options like Tangen, RPL etc." ,
      })
    }
  }
    
}
