// https://stackoverflow.com/questions/66571186
function setRangeDataValidation(range, list) {
  var rangelValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(list)
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rangelValidation);
}

function addBill() {
  let current = SpreadsheetApp.getActiveSpreadsheet();
  let templateSheet = current.getSheetByName("Template");
  let oldBill = current.getSheetByName("Bill");

  if(!templateSheet) throw new Error("Template sheet missing.");

  if(oldBill) {
    var confirm = Browser.msgBox('WARNING', "This will delete your old bill. Are you sure? (You won't be able to recover your data.)", Browser.Buttons.YES_NO); 
    if(confirm!='yes') return;

    current.deleteSheet(oldBill);
  }

  let sheetname = `Bill`;
  let newSheet = templateSheet.copyTo(current).setName(sheetname);

  let response = Browser.inputBox('Need Name', 'Enter bill holder name.', Browser.Buttons.OK);
  newSheet.getRange(2,3).setValue(response);
}

function addItem() {
  let current = SpreadsheetApp.getActiveSpreadsheet();
  let bill = current.getSheetByName("Bill");
  let items = current.getSheetByName("Items")

  if(!bill) return Browser.msgBox("ERROR", "There needs to be a sheet named \"Bill\" to add items.", Browser.Buttons.OK);
  if(!items) return Browser.msgBox("ERROR", "There needs to be a sheet named \"Items\" to add items.", Browser.Buttons.OK);

  let row = 0;

  for(let i = 7; i < 500; i++) {
    let cell = bill.getRange("A" + i);

    if(cell.getValue().toString() == "-") return Browser.msgBox("ERROR", "This bill has already been completed. Create a new bill.", Browser.Buttons.OK);

    if(cell.isBlank()) {
      row = i;
      break;
    }
  }

  if(row !== 0) {

      let codeValues = items.getRange("A2:A1000");

      let index = bill.getRange("A" + row);
      let itemCode = bill.getRange("B" + row);
      let itemName = bill.getRange("C" + row);
      let quantity = bill.getRange("D" + row);
      let unitPrice = bill.getRange("E" + row);
      let totalPrice = bill.getRange("F" + row);

      let combined = bill.getRange(`B${row}:F${row}`)

      index.setValue(row - 6);
      setRangeDataValidation(itemCode, codeValues)
      itemName.setValue(`=IFERROR(VLOOKUP($B$${row},Items!$A$2:$B,2,FALSE),"")`)
      quantity.setValue(1)
      unitPrice.setValue(`=IFERROR(VLOOKUP($B$${row},Items!$A$2:$C,3,FALSE),"")`)
      totalPrice.setValue(`=IF(OR(ISBLANK($D$${row}), ISBLANK($E$${row})), "", $D$${row}*$E$${row})`)

      if(row % 2 == 1) {
        combined.setBackground("#efefef")
      } else {
        combined.setBackground("#d9d9d9")
      }

      combined.setBorder(true, true, true, true, false, true, 'black', SpreadsheetApp.BorderStyle.SOLID)

      current.setActiveSheet(bill);
      current.setActiveRange(itemCode);
  }
}

function total() {

  let current = SpreadsheetApp.getActiveSpreadsheet();
  let bill = current.getSheetByName("Bill");
  let items = current.getSheetByName("Items")

  if(!bill) return Browser.msgBox("ERROR", "There needs to be a sheet named \"Bill\" to total values.", Browser.Buttons.OK);
  if(!items) return Browser.msgBox("ERROR", "There needs to be a sheet named \"Items\" to total values.", Browser.Buttons.OK);

  let row = 0;

  for(let i = 7; i < 500; i++) {
    let cell = bill.getRange("A" + i);

    if(cell.getValue().toString() == "-") return Browser.msgBox("ERROR", "This bill has already been completed. Create a new bill.", Browser.Buttons.OK);

    if(cell.isBlank()) {
      row = i;
      break;
    }
  }

  if(row == 7) return Browser.msgBox("ERROR", "You haven't added any items yet.", Browser.Buttons.OK);

  var confirm = Browser.msgBox('WARNING', "This action is irreversible. Do it anyway?", Browser.Buttons.YES_NO); 
  if(confirm!='yes') return;

  bill.setName(`bill_${Date.now()}`);

  if(row !== 0) {
      let codeValues = items.getRange("A2:A1000");

      let index = bill.getRange("A" + row);
      index.setValue("-");

      let all = bill.getRange(`E${row}:F${row+2}`);

      let subtotalText = all.getCell(1,1);
      let subtotalValue = all.getCell(1,2);

      let kdvText = all.getCell(2,1);
      let kdvValue = all.getCell(2,2);

      let totalText = all.getCell(3,1);
      let totalValue = all.getCell(3,2);

      all.setHorizontalAlignment('center');

      subtotalText.setBackground('#b7b7b7');
      kdvText.setBackground('#b7b7b7');
      totalText.setBackground('#b7b7b7');

      subtotalText.setValue("Subtotal:")
      kdvText.setValue(`=CONCATENATE("KDV (" , TEXT(Items!$E$2, 0), "%):")`)
      totalText.setValue(`Total:`)

      subtotalValue.setValue(`=SUM($F$7:$F${row-1})`);
      kdvValue.setValue(`=(Items!$E$2/100)*$F$${row}`);
      totalValue.setValue(`=SUM($F$${row},$F$${row+1})`);

      all.setBorder(true, true, true, true, true, true);

      SpreadsheetApp.setActiveSheet(bill);
      Browser.msgBox("Complete", "You can click the printer button on top left (in PC) in order to turn this page into a PDF.", Browser.Buttons.OK );
  }
}
