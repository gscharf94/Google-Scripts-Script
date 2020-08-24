// these are the hex values for the codes
// if you want to change the hue or whatever of the red for example
// just change the red to some other hex code
// search google for a hex color calculator or something like that
// you'll be able to find a hex code for w.e color you want
// there's literally millions of choices lol
const COLORS = {
  'darkGray':'#7c7c7c',
  'gray':'#d7d7d7',
  'red':'#ff2525',
  'lightRed':'#f58080',
  'orange':'#ffc622',
  'yellow':'#fdf322',
  'green':'#08bd0e',
};

function onOpen() {
  // this runs when the script is opened
  // it creates the Scripts menu option
  // for the user to run the script
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts').addItem('Call results','startSheet')
  .addItem('Callers Details', 'startCallerDetails')
  .addToUi();
}

function justToSaveStuff() {
  // this function solely exists so that my code editor
  // recognizes "Browser.msgBox" as a function
  // can delete this it doesn't matter at all
  Browser.msgBox('test');
}

function startSheet() {
  // entry point into the script
  let dict = compileRows();

  let timeInfo = {};
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.renameActiveSheet('Raw Data');

  // loops through the dictionary-like object
  // and creates an individual sheet for every canv
  for (const name in dict) {
    ss.insertSheet(name);
    let timeArr = populateIndividual(name, dict[name]);
    timeInfo[name] = {}
    timeInfo[name]['timeDiffs'] = timeArr[0];
    timeInfo[name]['startTime'] = timeArr[1];
    timeInfo[name]['endTime'] = timeArr[2];
  }
  populateGroup(dict, timeInfo);
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

function populateGroup(dict, timeInfo) {
  // creates the overview page and populates it with data
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet('Overview');
  let sheet = SpreadsheetApp.getActive().getSheetByName('Overview');
  ss.moveActiveSheet(1);
  sheet.setHiddenGridlines(true);
  
  const LETTERS = ['NULL','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
  
  function getUniqueResults() {
    // finds all the unique result options
    // and creates a template, which is used 
    // for the rest of the process
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


  // this part creates the header cells
  let topRow = ['Caller ID','Total Calls']
  for (const result in template) {
    topRow.push(result);
    topRow.push(`${result.slice(0,4)} Avg`);
  }
  topRow.push('Start Time');
  topRow.push('End Time');
  topRow.push('Hours Worked');
  topRow.push('5+ Time Diffs');
  let topRowRange = sheet.getRange(`A1:${LETTERS[topRow.length]}1`);
  let topRowRange2 = sheet.getRange(`B1:${LETTERS[topRow.length]}1`);
  topRowRange.setValues([topRow]);
  topRowRange2.setFontSize(11);

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

  // creates the meat of the page
  // tallies things up and writes it to the sheet
  let groupRows = []
  let c = 0;
  for (const name in dict) {
    let individualResults = getIndividualTemplate(name);
    let row = [name, dict[name].length];
    for (const result in individualResults) {
      row.push(individualResults[result]);
      row.push(individualResults[result]/dict[name].length);
    }
    groupRows.push(row);
  }


  let endCol = LETTERS[groupRows[0].length]
  let overviewRange = sheet.getRange(`A2:${endCol}${groupRows.length+1}`);
  overviewRange.setValues(groupRows);

  addSumTotals();
  let currentCol = 4;
  for (const result in template) {
    let range = sheet.getRange(`${LETTERS[currentCol]}2:${LETTERS[currentCol]}${groupRows.length+2}`);
    range.setNumberFormat('00.0%');
    range.setFontSize(11);
    currentCol += 2;
  }

  function addSumTotals() {
    // adds averages / totals to the bottom
    // and styles it appropriately
    // this part uses Sheets formulas because
    // I felt it would have been messier
    // to do it myself
    let row = groupRows.length+1;
    let toWrite = [['Averages'],['Totals']];

    toWrite[0].push(`=AVERAGE(B2:B${row})`);
    toWrite[1].push(`=SUM(B2:B${row})`);

    let col = 3;
    for (const result in template) {
      toWrite[0].push(`=AVERAGE(${LETTERS[col]}2:${LETTERS[col]}${row})`);
      toWrite[1].push(`=SUM(${LETTERS[col]}2:${LETTERS[col]}${row})`);
      toWrite[0].push(`=AVERAGE(${LETTERS[col+1]}2:${LETTERS[col+1]}${row})`);
      toWrite[1].push('');
      col += 2;
    }

    let range = sheet.getRange(`A${row+1}:${LETTERS[col-1]}${row+2}`);
    let range2 = sheet.getRange(`B${row+1}:${LETTERS[col-1]}${row+2}`);
    range.setValues(toWrite);
    range.setNumberFormat('####');
    range.setFontWeight('bold');
    range.setFontSize(11);

    range2.setBackground(COLORS['darkGray']);
  }

  function resizeColsRows() {
    // like the name says...
    // it resizes columns and rows
    // the values to the RIGHT of the functions
    // ie. setColumnWidth(i, HERE)
    // are changable, and that's the actual weight
    // 1 = A, 2 = B, ... etc
    sheet.setColumnWidth(1, 143);
    let i;
    for(i=2; i<topRow.length+1; i++) {
      sheet.setColumnWidth(i, 73);
    }
    sheet.setRowHeight(1, 57);
  }

  resizeColsRows();

  function formatStatic() {
    // formats the parts of the page that aren't dynamic
    let topRowRange = sheet.getRange(`A1:${LETTERS[topRow.length]}1`);
    let topRowRange2 = sheet.getRange(`B1:${LETTERS[topRow.length]}1`);
    topRowRange.setFontWeight('bold');
    topRowRange.setHorizontalAlignment('center');
    topRowRange.setWrap(true);

    topRowRange2.setBackground(COLORS['darkGray']);

    let leftColRange = sheet.getRange(`A1:A${groupRows.length+3}`);
    leftColRange.setFontWeight('bold');
    leftColRange.setHorizontalAlignment('right');
  }

  formatStatic();

  function addBorders() {
    // adds borders to the necessary cell ranges
    let leftRange = sheet.getRange(`A2:A${groupRows.length+3}`);
    leftRange.setBorder(null, null, null, true, false, false);
    let nextRange = sheet.getRange(`B2:B${groupRows.length+3}`);
    nextRange.setBorder(null, null, null, true, false, false);

    let col = 4;
    for (const result in template) {
      let lett = LETTERS[col];
      let range = sheet.getRange(`${lett}2:${lett}${groupRows.length+3}`);
      range.setBorder(null, null, null, true, false, false);
      col += 2;
    }
  }
  addBorders();

  function addStartEndTimes() {
    // adds the time info to the right of the avg stats
    let toWrite = [];

    function countTimeDiffs(name) {
      // counts the time diffs based on THRESHOLD value
      // to make it so there's a different threshold (10 mins instead of 5, for example)
      // just change THRESHOLD to whatever
      //               vvvv
      const THRESHOLD = 5;
      let rawData = timeInfo[name]['timeDiffs'];
      let count = 0;

      function loop(val, ind, arr) {
        let num = Number(String(val).split(" ")[0]);
        if(num >= THRESHOLD) {
          count++;
        }
      }
      rawData.forEach(loop);
      return count;
    }

    for (const name in dict) {
      let startTime = timeInfo[name]['startTime'];
      let endTime = timeInfo[name]['endTime'];
      let hoursWorked = new Date('1970/01/01 ' + endTime) - new Date ('1970/01/01 ' + startTime);
      hoursWorked = Math.round(hoursWorked/3600000 * 10) / 10;
      let timeDiffCount = countTimeDiffs(name);

      let row = [startTime, endTime, hoursWorked, timeDiffCount];
      toWrite.push(row);


    }

    let col = 4;
    for (const result in template) {
      col += 2;
    }
    let range = sheet.getRange(`${LETTERS[col-1]}2:${LETTERS[col+2]}${groupRows.length+1}`);
    sheet.setColumnWidth(col-1, 92);
    sheet.setColumnWidth(col, 92);
    sheet.setColumnWidth(col+1, 61);
    sheet.setColumnWidth(col+2, 61);

    let rightRange = sheet.getRange(`${LETTERS[col+2]}2:${LETTERS[col+2]}${groupRows.length+1}`);
    rightRange.setBorder(null, null, null, true, false, false);
    let bottomRange = sheet.getRange(`${LETTERS[col-1]}${groupRows.length+1}:${LETTERS[col+2]}${groupRows.length+1}`);
    bottomRange.setBorder(null, null, true, null, false, false);

    range.setValues(toWrite);
  }

  addStartEndTimes();

  function addAlternatingColors() {
    // makes it look prettier by making every other row
    // light gray
    let startRow = 2;
    let endRow = groupRows.length+1;

    let startCol = 2;
    let endCol = 6;
    for (const result in template) {
      endCol += 2;
    }

    let colorList = [];
    let i, j;
    for(i=startRow; i<endRow+1; i++) {
      row = [];
      for(j=startCol; j<endCol+1; j++) {
        if(i%2 == 0) {
          row.push('white');
        } else {
          row.push(COLORS['gray']);
        }
      }
      colorList.push(row);
    }

    let range = sheet.getRange(`${LETTERS[startCol]}${startRow}:${LETTERS[endCol]}${endRow}`);
    range.setBackgrounds(colorList);
  }
  addAlternatingColors();
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

  let rangePlusTimeDiffs = sheet.getRange(`A2:G${numRows+1}`);

  let i, j;
  let colorList = [];
  for(i=2; i<numRows+2; i++) {
    let row = [];
    for(j=1; j<8; j++) {
      if(i%2 == 0) {
        row.push('white');
      } else {
        row.push(COLORS['gray']);
      }
    }
    colorList.push(row);
  }

  rangePlusTimeDiffs.setBackgrounds(colorList);


  function addTimeDiffs() {
    // adds the time diffs next to the 'call time' column
    // returns an array [ startTime, endTime ]
    // so that the parent function can have it

    if(data.length == 1) {
      // if there is only one time, no use in time diffs
      let range = sheet.getRange('F2');
      return [[0],range.getValues(),range.getValues()];
    }
    let timeRange = sheet.getRange(`F2:F${data.length+1}`);
    let vals = timeRange.getValues();

    let timeDiffs = [];
    let avg = 0;

    function loop(cur, ind, arr) {
      if(ind == arr.length-1) {
        return;
      } else {
        let next = new Date ('1970/01/01 ' + arr[ind+1]);
        cur = new Date('1970/01/01 ' + cur);
        let timeDiff = next - cur;
        avg += timeDiff/60000;
        timeDiffs.push([`${timeDiff/60000} mins`]);
      }
    }    
    vals.forEach(loop);

    let diffRange = sheet.getRange(`G2:G${timeDiffs.length+1}`);
    diffRange.setValues(timeDiffs);

    let startTime = vals[0];
    let endTime = vals[vals.length-1];

    avg = Math.round((avg/timeDiffs.length+1)*100)/100;
    let avgRange = sheet.getRange(`H2:I2`);
    let writeTo = ['Avg Time Diff', `${avg} mins`]
    let weights = ['bold','normal'];

    avgRange.setValues([writeTo]);
    avgRange.setFontWeights([weights]);

    return [timeDiffs, startTime, endTime];
  }
  timeArr = addTimeDiffs();

  function formatStatic() {
    // formats some of the things that won't change
    sheet.setColumnWidth(1, 77);
    sheet.setColumnWidth(2, 131);
    sheet.setColumnWidth(3, 93);
    sheet.setColumnWidth(4, 65);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 92);
    sheet.setColumnWidth(7, 64);

    let topRow = sheet.getRange(`A1:G1`);
    topRow.setFontWeight('bold');
    topRow.setBackground(COLORS['darkGray']);
    sheet.setHiddenGridlines(true)

    let bottomRow = sheet.getRange(`A${numRows+1}:G${numRows+1}`);
    bottomRow.setBorder(null, null, true, null, false, false);

  }

  formatStatic();

  function colorTimeDiffs() {
    // this adds borders to the rightmost column
    // as well as color codes the time diffs column
    let range = sheet.getRange(`G2:G${numRows+1}`);
    let vals = range.getValues();

    colors = [];
    weights = [];

    c = 2;
    // if you want to change the color scheme
    // just change the numbers 25, 0, 2
    // right now, if it's bigger than 25, it gets red
    // if it's less than 25 but bigger than 2, it gets orange
    // if it's 0 it gets yellow
    // the rest, no change
    // you can freely edit those values
    function loop(val, ind, arr) {
      let num = Number(String(val).split(" ")[0]);
      // here vvvv
      if(num >= 25) {
        colors.push(['red']);
        weights.push(['bold']);
      //       here vvvv
      } else if(num == 0) {
        colors.push(['yellow']);
        weights.push(['bold']);
      //        here vvvv 
      } else if(num > 2) {
        colors.push(['orange']);
        weights.push(['bold']);
      } else {
        if(c%2 == 0) {
          colors.push(['white']);
          weights.push(['normal']);
        } else {
          colors.push([COLORS['gray']]);
          weights.push(['normal']);
        }
      }
      c++;
    }
    vals.forEach(loop);
    range.setBackgrounds(colors);
    range.setFontWeights(weights);
    range.setBorder(null, null, null, true, false, false);
  }

  colorTimeDiffs();
  return timeArr;
}

function getRangeVals() {
  // helper function to help get the actual
  // text data from a range of cells
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
    let phone = String(val[4]);
    phone = `(${phone.slice(0,3)}) ${phone.slice(3,6)}-${phone.slice(6,10)}`;
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

function determineMessage(message) {
  // takes in a message and figures out if
  // A) it has the word "wrong" or "equivocado" in it
  // B) it has the word "stop" in it
  // C) there was a call back

  let wrong, stop, call;

  message = message.toLowerCase();
  if(message.indexOf('wrong') == -1 || message.indexOf('equivocado') == -1) {
    wrong = false;
  } else {
    wrong = true;
  }

  if(message.indexOf('stop') == -1) {
    stop = false;
  } else {
    stop = true;
  }

  if(message.indexOf('tried to call you.') == -1) {
    call = false;
  } else {
    call = true;
  }

  return [wrong, stop, call];
}

function startText() {
  
  function getDataRange() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();
    let range = sheet.getDataRange();
    return range.getValues();
  }
  
  function organizeIntoObject(rawData) {
    
    let dict = {};
    
    function addNames(val, ind, arr) {
      let canvName = `${val[4]} ${val[5]}`;
      if(canvName == "sender_first_name sender_last_name" || canvName == "") {
        return;
      } else {
        dict[canvName] = {
          'total':0,
          'wrong#':0,
          'stop':0,
          'tried to call':0,
        };
      }
    }
    rawData.forEach(addNames);

    let i;
    let current = null;
    for(i=1; i<rawData.length; i++) {
      let status = rawData[i][2];
      let message = rawData[i][3];
      let canvName = `${rawData[i][4]} ${rawData[i][5]}`;

      if(status == "outgoing") {
        if(current == null) {

        } else {
          dict[canvName]['total']++;
          let result = determineMessage(current['message']);
          if(result[0]) {
            dict[canvName]['wrong#']++;
          }
          if(result[1]) {
            dict[canvName]['stop']++;
          }
          if(result[2]) {
            dict[canvName]['tried to call']++;
          }
        }
        current = {
          'name':canvName,
          'message':message,
        };
      } else {
        current['message'] += message;
      }
    }

    for (const name in dict) {
      let str1 = ` | total: ${dict[name]['total']} `;
      let str2 = `wrong: ${dict[name]['wrong#']} `;
      let str3 = `stop: ${dict[name]['stop']} `;
      let str4 = `tried to call: ${dict[name]['tried to call']}`;
      Browser.msgBox(name+str1+str2+str3+str4);
    }    
  }
  
  let rawData = getDataRange();
  let processedData = organizeIntoObject(rawData);
}

function startCallerDetails() {
  
  let rangeVals = getDataRange();
  let dict = createDict(rangeVals);
  
  setupSheet();
  
  function getDataRange() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();
    let range = sheet.getDataRange();
    return range.getValues();
  }
  
  function createDict(rawData) {
    let dict = {};
    function addNames(val, ind, arr) {
      if(ind == 0) {
        return;
      }
      dict[val[2]] = {
        'login':val[1],
        'email':val[3],
        'inCall':val[5],
        'inWrap':val[6],
        'inReady':val[7],
        'inNotReady':val[8],
        'totalCalls':val[9],
      };
    }
    rawData.forEach(addNames);
    return dict;
  }
    
  function setupSheet() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.renameActiveSheet('Raw Data');
    ss.insertSheet("Formatted");
    let sheet = SpreadsheetApp.getActive().getSheetByName('Formatted');
    sheet.setHiddenGridlines(true);
    
    let topRow = ['Login','Name','Email','Wrap Up','Not Ready','Call','Ready','Total'];
    
    let topRange = sheet.getRange('A1:H1');
    topRange.setValues([topRow]);
    topRange.setFontWeight('bold');
    topRange.setFontSize(11);
    topRange.setBackground(COLORS['darkGray']);
    topRange.setHorizontalAlignment('center');

    }
    



  function createToWriteArr() {
    let toWrite = [];
    for (const name in dict) {
      
      let row = [
        dict[name]['login'],
        name,
        dict[name]['email'],
        dict[name]['inWrap'],
        dict[name]['inNotReady'],
        dict[name]['inCall'],
        dict[name]['inReady'],
        dict[name]['totalCalls'],
      ];

      toWrite.push(row);
    }
    return toWrite;
  }

  let unsortedData = createToWriteArr();

  let sortedData = unsortedData.sort(function(a,b) {
    let nA = a[1].toLowerCase().charCodeAt();
    let nB = b[1].toLowerCase().charCodeAt();
    return nA - nB;
  });

  function createColorArr(width, height) {
    let arr = [];
    for(let i=0; i<height; i++) {
      let row = [];
      for(let j=0; j<width; j++) {
        if(i%2 == 0) {
          row.push('white');
        } else {
          row.push(COLORS['gray']);
        }
      }
      arr.push(row);
    }
    return arr;
  }

  function writeToSheet(sData) {
    let sheet = SpreadsheetApp.getActive().getSheetByName('Formatted');
    let eRow = 1 + sData.length;
    let range = sheet.getRange(`A2:H${eRow}`);
    range.setValues(sData);

    let colors = createColorArr(8, sData.length);
    range.setBackgrounds(colors);

    sheet.setColumnWidth(1,108);
    sheet.setColumnWidth(2,177);
    sheet.setColumnWidth(3,177);
    sheet.setColumnWidth(4,74);
    sheet.setColumnWidth(5,74);
    sheet.setColumnWidth(6,74);
    sheet.setColumnWidth(7,74);
    sheet.setColumnWidth(8,74);
    sheet.setColumnWidth(10,200);
    sheet.setColumnWidth(11,70);
    sheet.setColumnWidth(12,70);
    sheet.setColumnWidth(14,200);

    let avgRowRange = sheet.getRange(`A${eRow+1}:H${eRow+1}`);
    let avgRow = [
      '',
      '',
      'AVERAGE',
      `=AVERAGE(D2:D${eRow})`,
      `=AVERAGE(E2:E${eRow})`,
      `=AVERAGE(F2:F${eRow})`,
      `=AVERAGE(G2:G${eRow})`,
      `=AVERAGE(H2:H${eRow})`,
    ];
    avgRowRange.setValues([avgRow]);
    avgRowRange.setFontSize(11);
    avgRowRange.setFontWeight('bold');
    avgRowRange.setBackground(COLORS['darkGray']);
    avgRowRange.setHorizontalAlignment('right');
    avgRowRange.setNumberFormat('###.#');

    sheet.getRange(`B${eRow+1}:C${eRow+1}`).merge();

    sheet.getRange('K3:L3').merge();
    sheet.getRange('K4:L4').merge();
    sheet.getRange('K5:L5').merge();
    sheet.getRange('K6:L6').merge();
    sheet.getRange('K7:L7').merge();
    sheet.getRange('K3').setValue('Minutes in Wrap Up');
    sheet.getRange('K3').setHorizontalAlignment('center');
    sheet.getRange('K4').setValue('Minutes in Not Ready');
    sheet.getRange('K4').setHorizontalAlignment('center');
    sheet.getRange('K5').setValue('Minutes in Call');
    sheet.getRange('K5').setHorizontalAlignment('center');
    sheet.getRange('K6').setValue('Minutes in Ready');
    sheet.getRange('K6').setHorizontalAlignment('center');
    sheet.getRange('K7').setValue('Total Calls');
    sheet.getRange('K7').setHorizontalAlignment('center');

    let leftCol = [
      [`=VLOOKUP($J$2,$B$2:$H$${eRow+1},3,FALSE)`],
      [`=VLOOKUP($J$2,$B$2:$H$${eRow+1},4,FALSE)`],
      [`=VLOOKUP($J$2,$B$2:$H$${eRow+1},5,FALSE)`],
      [`=VLOOKUP($J$2,$B$2:$H$${eRow+1},6,FALSE)`],
      [`=VLOOKUP($J$2,$B$2:$H$${eRow+1},7,FALSE)`],
    ];
    
    let rightCol = [
      [`=VLOOKUP($M$2,$B$2:$H$${eRow+1},3,FALSE)`],
      [`=VLOOKUP($M$2,$B$2:$H$${eRow+1},4,FALSE)`],
      [`=VLOOKUP($M$2,$B$2:$H$${eRow+1},5,FALSE)`],
      [`=VLOOKUP($M$2,$B$2:$H$${eRow+1},6,FALSE)`],
      [`=VLOOKUP($M$2,$B$2:$H$${eRow+1},7,FALSE)`],
    ];

    let leftRange = sheet.getRange(`J3:J7`);
    leftRange.setValues(leftCol);

    let rightRange = sheet.getRange('M3:M7');
    rightRange.setValues(rightCol);
    rightRange.setHorizontalAlignment('left');

    let names = sheet.getRange(`B2:B${eRow+1}`);
    let rule = SpreadsheetApp.newDataValidation().requireValueInRange(names).build();

    let leftCell = sheet.getRange('J2');
    leftCell.setDataValidation(rule);
    leftCell.setValue(sheet.getRange('B2').getValue());
    leftCell.setFontWeight('bold');
    let rightCell = sheet.getRange('M2');
    rightCell.setDataValidation(rule);
    rightCell.setValue('AVERAGE');
    rightCell.setFontWeight('bold');
  }

  writeToSheet(sortedData);

  function createFirstChart() {
    let sheet = SpreadsheetApp.getActive().getSheetByName('Formatted');
    let chartBuilder = sheet.newChart();
    let leftRange = sheet.getRange('J2:M7');
    let rightRange = sheet.getRange('M2:M7');
    chartBuilder.addRange(leftRange)
    .setChartType(Charts.ChartType.BAR)
    .setPosition(8,9,0,0);
    // chartBuilder.addRange(rightRange);

    sheet.insertChart(chartBuilder.build());
  }

  function writeSecondPart() {
    let sheet = SpreadsheetApp.getActive().getSheetByName('Formatted');

    let cell = sheet.getRange('J28');
    cell.setValue('=J2&" vs "&M2');
    cell.setFontWeight('bold');

    let toWrite = [
      [
        '=J3-M3',
        '=IF(J29<0, "less than", "more than")',
        '=M2',
        'Minutes in Wrap Up',
      ],
      [
        '=J4-M4',
        '=IF(J30<0, "less than", "more than")',
        '=M2',
        'Minutes in Not Ready',
      ],
      [
        '=J5-M5',
        '=IF(J31<0, "less than", "more than")',
        '=M2',
        'Minutes in Call',
      ],
      [
        '=J6-M6',
        '=IF(J32<0, "less than", "more than")',
        '=M2',
        'Minutes in Ready',
      ],
      [
        '=J7-M7',
        '=IF(J32<0, "less than", "more than")',
        '=M2',
        'Total Calls',
      ]
    ];

    let range = sheet.getRange('J29:M33');
    range.setValues(toWrite);


  }

  createFirstChart();
  writeSecondPart();

  function createSecondChart() {
    let sheet = SpreadsheetApp.getActive().getSheetByName('Formatted');
    let chartBuilder = sheet.newChart();
    let leftRange = sheet.getRange('J29:J33');
    // chartBuilder.addRange(leftRange)
    // .setChartType(Charts.ChartType.COLUMN)
    // .setPosition(34,9,0,0);


    chartBuilder.addRange(sheet.getRange('J29'))
    .addRange(sheet.getRange('J30'))
    .addRange(sheet.getRange('J31'))
    .addRange(sheet.getRange('J32'))
    .addRange(sheet.getRange('J33'))
    .setChartType(Charts.ChartType.COLUMN)
    .setPosition(34,9,0,0);


    let test = chartBuilder.asColumnChart();
    
    let vals = leftRange.getValues();

    let colors = [];
    vals.forEach(
      (val, ind, arr) => {
        Browser.msgBox(`val: ${val} type: ${typeof(val)}`);
        Browser.msgBox(`val[0]: ${val[0]} type: ${typeof(val[0])}`);
        if(val > 0) {
          colors.push(['green']);
        } else {
          colors.push(['red']);
        }
      }
    )
    
    test.setColors(colors);



    sheet.insertChart(test.build());
  }

  createSecondChart();

}

