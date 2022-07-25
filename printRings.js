function printScoresheets(level = "Beginner") {
  // Print the combined forms/sparring scoresheet for particular level

  // 
  var targetDocName = level + " Forms Score Sheets"

  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(level)

  var targetDoc = createDocFile(targetDocName)
  targetDoc.setText('')

  var targetBody = targetDoc.getBody()
  var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet)

  var physToVirtMap = physToVirtMapInv(virtToPhysMap)

  // Iterate through the list of sorted physical rings
  // For now, use a new spreadsheet.

  var targetSheetName = level + " Sparring Score Sheets"
  var targetSpreadsheet = createSpreadsheetFile(targetSheetName)


  for (var physRingStr of sortedPhysRings(virtToPhysMap)) {
    var virtRing = physToVirtMap[physRingStr]


    // Get all the people in one virtRing, whether forms, sparring, or both
    var virtRingPeople = peopleArr
      .filter((person) => person.vRing == virtRing)

    // Make the forms scoresheet
    // filter on doing forms and then sort
    var formsPeople = virtRingPeople.filter((person) => person.forms != "no")
    .sort(sortByFormOrder)

    appendOneFormsScoresheet(targetBody, formsPeople, virtRing, physRingStr, level)
    console.log('Finished with forms ring ' + physRingStr)
  
    // Make the sparring scoresheet
    // filter on doing forms and then sort
    var sparringPeople = virtRingPeople.filter((person) => person.sparring != "no")
    .sort(sortBySparringOrder)

    appendOneSparringScoresheet(targetSpreadsheet, sparringPeople, virtRing, physRingStr, level)

    console.log('Finished with sparring ring ' + physRingStr)
  }

  targetDoc.saveAndClose()


}

// Create the ring sheet for one level.
function printRingsOneLevel(
  sourceSheetName,
  readFromCalcRings = false,
  useRemapping = false
) {
  // sourceSheetName must be the sheet name for one of the levels.
  // "sourceSheetName Rings" must be another existing sheet. The target will be cleared each time.

  var targetSheetName = sourceSheetName + " Rings";
  var targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(sourceSheetName);

  // Clear the target to redo form
  targetSheet.clear();

  // peopleArr is going to be the student data read into an hash of array of hashes.
  var [peopleArr, virtToPhysMap] = readTableIntoArr(
    sourceSheet,
    readFromCalcRings
  );

  var phyRingColorMap = globalVariables().phyRingColorMap;

  // Need to get all the vrings here

  var x;
  var y;
  for (var vRing of Object.keys(virtToPhysMap)) {
    // convert virtual to physical
    var physRingStr = virtToPhysMap[vRing].toString();
    var physArr = physRingStr.match(/\d+|\D+/g);

    x = parseInt(physArr[0]) - 1;
    var physRingNumber = x + 1;
    if (physArr[1] == "b") {
      y = 1;
    } else {
      y = 0;
    }
    // x and y are 0-based indices into the table

    var startCol = 1 + 7 * x;
    var startRow = 1 + 25 * y;

    var peopleInThisVRing = peopleArr.filter((person) => person.vRing == vRing);
    printRing(
      targetSheet,
      startCol,
      startRow,
      vRing,
      peopleInThisVRing,
      physRingStr,
      phyRingColorMap[physRingNumber]
    );
  }
}

function printRing(
  targetSheet,
  startCol,
  startRow,
  ringId,
  peopleArr,
  phyRing,
  phyRingColor = null
) {
  var mainHeaderRows = printMainHeader(
    targetSheet,
    startRow,
    startCol,
    ringId,
    phyRing,
    phyRingColor
  );

  // Get an array of formers
  var formerArr = [];
  for (var i = 0; i < peopleArr.length; i++) {
    if (peopleArr[i].form != "No") {
      formerArr.push(peopleArr[i]);
    }
  }
  var formHeaderRows = printFormsHeader(
    targetSheet,
    startRow,
    startCol,
    formerArr.length
  );
  var formStartRow = startRow + mainHeaderRows + formHeaderRows;

  var formRows = printPeopleArr(targetSheet, formerArr, formStartRow, startCol);

  // Get an array of sparrers
  var sparrerArr = [];
  for (var i = 0; i < peopleArr.length; i++) {
    if (peopleArr[i].sparring != "No") {
      sparrerArr.push(peopleArr[i]);
    }
  }
  var numSparrers = sparrerArr.length;

  var sparHeaderRows = printSparHeader(
    targetSheet,
    startRow + mainHeaderRows + formHeaderRows + formRows,
    startCol,
    numSparrers
  );
  var sparStartRow = startRow + mainHeaderRows + formHeaderRows + formRows;

  var numSparrers = printPeopleArr(
    targetSheet,
    sparrerArr,
    sparStartRow + sparHeaderRows,
    startCol
  );

  // Border all the way around
  var cells = targetSheet.getRange(
    startRow,
    startCol,
    2 + formerArr.length + numSparrers + 3,
    7
  );
  cells.setBorder(
    true,
    true,
    true,
    true,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.SOLID_THICK
  );

  // After the Ring heading
  cells = targetSheet.getRange(startRow + 1, startCol, 1, 7);
  cells.setBorder(
    null,
    null,
    true,
    null,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.DOUBLE
  );

  // Border after the main heading
  cells = targetSheet.getRange(startRow + 2, startCol, 1, 7);
  cells.setBorder(null, null, true, null, null, null);

  // Border after forms section
  var cells = targetSheet.getRange(sparStartRow, startCol, 1, 7);
  cells.setBorder(null, null, true, null, null, null);

  targetSheet.autoResizeColumns(startCol, 7);
}

function printSparHeader(targetSheet, startRow, startCol, numSparrers) {
  targetSheet.getRange(startRow, startCol, 1, 2).setNumberFormat("@");
  targetSheet
    .getRange(startRow, startCol, 1, 2)
    .setValues([["Sparring", "(" + numSparrers + ")"]])
    .setFontSize(16)
    .setFontWeight("bold");
  return 1;
}

//  for (var k = 0; k<peopleArr.length; k++) {
//    if (peopleArr[k].sparring == "Yes") {
//      cell = targetSheet.getRange(startRow +1 + sparrerNum, startCol, 1, 7);
//      cell.setValues([[peopleArr[k].sfn, peopleArr[k].sln]]);
//      printPeopleArr(targetSheet, peopleArr, startRow+1+sparrerNum, startCol);
//     sparrerNum++;
//    }
//  }
// targetSheet.getRange(startRow, startCol + 1).setNumberFormat("@")
// targetSheet.getRange(startRow, startCol,1,2).setValues([["Sparring", "\(" + sparrerNum + "\)"]]).setFontSize(16).setFontWeight('bold');
//  return sparrerNum;
//}

function printMainHeader(
  targetSheet,
  startRow,
  startCol,
  ring,
  phyRing,
  phyRingColor
) {
  var cells = targetSheet.getRange(startRow, startCol);
  cells
    .setValue("Ring " + phyRing)
    .setFontSize(20)
    .setFontWeight("bold");

  // set background color
  cells = targetSheet.getRange(startRow, startCol, 1, 7);
  cells.setBackgroundColor(phyRingColor);
  // change font color if black
  if (["black", "#0000ff"].includes(phyRingColor)) {
    cells.setFontColor("white");
  }
  cells = targetSheet.getRange(startRow + 1, startCol);
  cells
    .setValue("(virtual ring " + ring + ")")
    .setFontSize(16)
    .setFontWeight("bold");

  return 2; // the number of rows printed
}

// Print out the header cells for one ring
function printFormsHeader(sheet, row, col, numForms) {
  const headers = [
    "First",
    "Last",
    "Age",
    "Height",
    "School",
    "Sparring?",
    "gender",
  ];
  sheet.getRange(row + 2, col, 1, 2).setNumberFormat("@");
  sheet
    .getRange(row + 2, col, 1, 2)
    .setValues([["Forms", "(" + numForms + ")"]])
    .setFontSize(16)
    .setFontWeight("bold")
    .setNumberFormat("@");
  for (let index = 0; index < headers.length; index++) {
    cells = sheet.getRange(row + 3, index + col);
    cells.setValue(headers[index]).setFontWeight("bold");
  }
  //cells.setValue("a");

  return 2; // rows in forms header
}

// Fill out the info for one ring. Just the data, not the headers.
function printPeopleArr(targetSheet, peopleArr, row, col) {
  const ringHeaders = [
    "sfn",
    "sln",
    "age",
    "height",
    "school",
    "sparring",
    "gender",
  ];
  var cells = [];
  for (let i = 0; i < peopleArr.length; i++) {
    cells.push([]);
    for (let j = 0; j < ringHeaders.length; j++) {
      cells[i].push(peopleArr[i][ringHeaders[j]]);
    }
  }
  targetSheet
    .getRange(row, col, peopleArr.length, ringHeaders.length)
    .setValues(cells)
    .setHorizontalAlignment("left");

  return peopleArr.length;
}
// Fill out the info for one ring. Just the data, not the headers.
function printPeopleArrOld(targetSheet, peopleArr, row, col) {
  const ringHeaders = [
    "sfn",
    "sln",
    "age",
    "height",
    "school",
    "sparring",
    "gender",
  ];
  for (let i = 0; i < peopleArr.length; i++) {
    for (let j = 0; j < ringHeaders.length; j++) {
      targetSheet
        .getRange(row + i, col + j)
        .setValue(peopleArr[i][ringHeaders[j]])
        .setHorizontalAlignment("left");
    }
  }

  return peopleArr.length;
}
