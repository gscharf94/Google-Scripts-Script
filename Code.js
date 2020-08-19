function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts').addItem('Process Sheet','startSheet').addToUi();
}

function justToSaveStuff() {
  Browser.msgBox('test');
}

function startSheet() {
  // entry point into the script
  let dict = compileRows();

  let namesList = [];
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('Raw Data');
  for (const name in dict) {
    ss.insertSheet(name);
    populateIndividual(name, dict[name]);
  }
  
  populateGroup(dict);
  
  // tmp = ['PoderFLVRP7']
  // createNewSheets(tmp);
  // populateIndividual('PoderFLVRP7',dict['PoderFLVRP7']);
}

function sortDictByTime(dict) {
  // sorts individual data by time
  // in descending order
  for (const name in dict) {
    let data = dict[name];
    let sortedData = data.sort(function(a,b) {
      return new Date('1970/01/01 ' + a[5]) - new Date('1970/01/01 ' + b[5]);
    });
    dict[name] = sortedData;
  }
  return dict;
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

function countObj(obj) {
  let c = 0;
  for (const thing in obj) {
    c++;
  }
  return c;
}

function populateGroup(dict) {
  // creates the overview page and populates it with data
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet('Overview');
  let sheet = SpreadsheetApp.getActive().getSheetByName('Overview');
  ss.moveActiveSheet(1);
  
  const LETTERS = ['NULL','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
  
  function getUniqueResults() {
    let template = {};
    for (const name in dict) {
      let data = dict[name];
      for(let i=0; i<data.length; i++) {
        let res = data[i][4];
        template[res] = 0;
      }
    }
    return template;
  }
  
  let template = getUniqueResults();
  
  let topRow = ['Caller ID','Total Calls']
  for (const result in template) {
    topRow.push(result);
  }
  let topRowRange = sheet.getRange(`A1:${LETTERS[topRow.length]}1`);
  topRowRange.setValues([topRow]);

  function getIndividualTemplate(name) {
    // counts up totals from results
    let templateCopy = { ...template };
    let data = dict[name];
    
    function loop(val, ind, arr) {
      let result = val[4];
      templateCopy[result]++;
    }

    data.forEach(loop);

    return templateCopy;
  }

  let groupRows = []
  for (const name in dict) {
    let individualResults = getIndividualTemplate(name);
    let row = [name, dict[name].length];
    for (const result in individualResults) {
      row.push(individualResults[result]);
    }
    groupRows.push(row);
  }


  let endCol = LETTERS[groupRows[0].length]
  let overviewRange = sheet.getRange(`A2:${endCol}${groupRows.length+1}`);
  overviewRange.setValues(groupRows);

  // Browser.msgBox(groupRows);

}

function populateIndividual(callerID, data) {
  // takes in callerID's data from dict
  // and populates their spreadsheet with their
  // individual data
  let sheet = SpreadsheetApp.getActive().getSheetByName(callerID);

  let topRow = [['Voter ID', 'Voter Name','Voter Phone','Call Date','Result','Call Time','Time Diff']];
  let topRowRange = sheet.getRange('A1:G1');
  topRowRange.setValues(topRow);

  let numRows = data.length;
  let range = sheet.getRange(`A2:F${numRows+1}`);
  range.setValues(data);

  function addTimeDiffs() {
    // adds the time diffs next to the 'call time' column
    // returns an array [ startTime, endTime ]
    // so that the parent function can have it

    if(data.length == 1) {
      // if there is only one time, no use in time diffs
      return;
    }

    let timeRange = sheet.getRange(`F2:F${data.length+1}`);
    let vals = timeRange.getValues();

    let timeDiffs = [];

    function loop(cur, ind, arr) {
      if(ind == arr.length-1) {
        return;
      } else {
        let next = new Date ('1970/01/01 ' + arr[ind+1]);
        cur = new Date('1970/01/01 ' + cur);
        let timeDiff = next - cur;
        timeDiffs.push([`${timeDiff/60000} mins`]);
      }
    }
    
    vals.forEach(loop);

    let diffRange = sheet.getRange(`G2:G${timeDiffs.length+1}`);
    diffRange.setValues(timeDiffs);

    let startTime = vals[0];
    let endTime = vals[vals.length-1];
    return [startTime, endTime];
  }

  startEndTimes = addTimeDiffs();
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
        date, result, time
      ]); 
    }
  }
  let dict = {};
  rawData.forEach(addNames);
  rawData.forEach(loop);

  dict = sortDictByTime(dict);

  return dict;
}