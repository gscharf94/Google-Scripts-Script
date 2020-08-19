function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts').addItem('Process Sheet','startSheet').addToUi();
}

function startSheet() {
  // entry point into the script
  let dict = compileRows();

  let namesList = [];
  for (const name in dict) {
    namesList.push(name);
  }
  createNewSheets(namesList);


}

function createNewSheets(namesList) {
  // takes in a list of caller ids
  // and creates new sheets for each of them
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  function createNewSheet(val, ind, arr) {
    ss.insertSheet(val);
  }
  namesList.forEach(createNewSheet);
}

function populateIndividual(callerID, data) {
  // takes in callerID's data from dict
  // and populates their spreadsheet with their
  // individual data
  let sheet = SpreadsheetApp.getActive().getSheetByName(callerID);

  let numRows = data.length;
  let range = sheet.getRange(`A2:G${numRows+1}`);

}

function getRangeVals() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let rangeData = sheet.getDataRange();
  return rangeData.getValues();
}

function compileRows() {
  // creates a python-dictionary basically structured like so
  // dict = {
  //  ${caller login}: [row, row, row],
  //  ${caller login}: [row, row, row],
  // }
  let rawData = getRangeVals();

  function addNames(val, ind, arr) {
    let callerID = val[7];
    if(callerID == "" || callerID == "Caller Login") {
      return;
    } else {
      dict[callerID] = [];
    }
  }

  function loop(val, ind, arr) {
    let voterID = val[0];
    let fullName = `${val[2]} ${val[3]}`;
    let phone = val[4];
    let date = val[5];
    let time = val[6];
    let callerID = val[7];
    if(callerID == "" || callerID == "Caller Login") {
      return;
    } else {
      let result = val[8];
      dict[callerID].push([
        voterID, fullName, phone,
        date, time, result
      ]); 
    }
  }
  let dict = {};
  rawData.forEach(addNames);
  rawData.forEach(loop);
  return dict;
}