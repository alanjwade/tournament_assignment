// Entry point. Do all the levels, or pick one to do.
function generateOverview(
  level = false,
  readFromCalcRings = false,
  useRemapping = false
) {
  const catArr = globalVariables().levels

  if (level) {
    generateOverviewOneLevel(
      level,
      readFromCalcRings,
      (useRemapping = useRemapping)
    )
  } else {
    for (var i = 0; i < catArr.length; i++) {
      generateOverviewOneLevel(
        catArr[i],
        readFromCalcRings,
        (useRemapping = useRemapping)
      )
    }
  }
}

function getPhysRingNumber(physRingStr) {
  var physArr = physRingStr.match(/\d+|\D+/g)

  var x = parseInt(physArr[0]) - 1
  var physRingNumber = x + 1

  return physRingNumber
}

// Create the ring sheet for one level.
function generateOverviewOneLevel(
  level,
  readFromCalcRings = false,
  useRemapping = false
) {
  // level must be the sheet name for one of the levels.
  // "level Rings" must be another existing sheet. The target will be cleared each time.

  var targetSheetName = level + " Rings"

  // Open or create a separate spreadsheet
  var targetSpreadsheet = openOrCreateFileInFolder(
    level + ' Overview', (isSpreadSheet = true)
  )

  // If the sheet doesn't exist make it
  if (targetSpreadsheet.getSheetByName(targetSheetName) == null) {
    targetSpreadsheet.insertSheet(targetSheetName)
  }
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName)


  // Remove all other sheets
  var allTargetSheets = targetSpreadsheet.getSheets()
  for (var i=0; i<allTargetSheets.length - 1; i++) {
    if (allTargetSheets[i].getName() != targetSheetName) {
      targetSpreadsheet.deleteSheet(allTargetSheets[i])
    }
  }

  //var targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName)
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(level)

  // Clear the target to redo form
  targetSheet.clear()

  // peopleArr is going to be the student data read into an hash of array of hashes.
  var [peopleArr, virtToPhysMap] = readTableIntoArr(
    sourceSheet,
    readFromCalcRings
  )

  // Need to get all the vrings here

  var x
  var y
  for (var vRing of Object.keys(virtToPhysMap)) {
    // convert virtual to physical
    var physRingStr = virtToPhysMap[vRing].toString()
    var physArr = physRingStr.match(/\d+|\D+/g)

    x = parseInt(physArr[0]) - 1
    if (physArr[1] == "b") {
      y = 1
    } else {
      y = 0
    }
    // x and y are 0-based indices into the table

    var startCol = 1 + 7 * x
    var startRow = 1 + 25 * y

    var peopleInThisVRing = peopleArr.filter((person) => person.vRing == vRing)
    generateOverviewOneRing(
      targetSheet,
      startCol,
      startRow,
      vRing,
      peopleInThisVRing,
      physRingStr
    )

    // Generate a timestamp
    targetSheet.getRange(51, 1).setValue(createTimeStamp())
    targetSheet.getRange(51, 1, 1, 5).mergeAcross()
  }
}

function generateOverviewOneRing(
  targetSheet,
  startCol,
  startRow,
  ringId,
  peopleArr,
  physRing
) {
  

  // ************************************
  // main header
  // ************************************
  var mainHeaderRows = printMainHeader(
    targetSheet,
    startRow,
    startCol,
    ringId,
    physRing,
    numCols = 7
  )

  var mainHeaderRow = startRow
  var curRow = startRow + mainHeaderRows

  // Get an array of formers
  var formerArr = []
  for (var i = 0; i < peopleArr.length; i++) {
    if (peopleArr[i].form.toLowerCase() != "no") {
      formerArr.push(peopleArr[i])
    }
  }
  formerArr = formerArr.sort(sortByFormOrder)

  // **********************************
  // form header
  // **********************************
  var formHeaderRows = printFormsHeader(
    targetSheet,
    curRow,
    startCol,
    formerArr.length,
    numCols
  )

  var formHeaderRow = curRow

  curRow += formHeaderRows

  curRow += printGeneralHeader(sheet, curRow, col)


  // *********************************
  // Array of form people
  // *********************************
  var formRows = printPeopleArr(targetSheet, formerArr, curRow, startCol)

  curRow += formRows

  // Get an array of sparrers
  var sparrerArr = []
  for (var i = 0; i < peopleArr.length; i++) {
    if (peopleArr[i].sparring.toLowerCase() != "no") {
      sparrerArr.push(peopleArr[i])
    }
  }
  sparrerArr = sparrerArr.sort(sortBySparringOrder)
  var numSparrers = sparrerArr.length

  // *********************************
  // Sparring header
  // *********************************
  var sparHeaderRows = printSparHeader(
    targetSheet,
    curRow,
    startCol,
    numSparrers
  )

  var sparHeaderRow = curRow
  curRow += sparHeaderRows
  curRow += printGeneralHeader(sheet, curRow, col)

  // *******************************
  // Sparring array
  // *******************************
  var numSparrers = printPeopleArr(
    targetSheet,
    sparrerArr,
    curRow,
    startCol
  )

  curRow += numSparrers

  // Border all the way around
  var cells = targetSheet.getRange(
    startRow,
    startCol,
    2 + formerArr.length + numSparrers + 3,
    numCols
  )
  cells.setBorder(
    true,
    true,
    true,
    true,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.SOLID_THICK
  )

  // After the Ring heading
  cells = targetSheet.getRange(startRow + 1, startCol, 1, numCols)
  cells.setBorder(
    null,
    null,
    true,
    null,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.DOUBLE
  )

  // Border after the main heading
  cells = targetSheet.getRange(startRow + 2, startCol, 1, numCols)
  cells.setBorder(null, null, true, null, null, null)

  // Border after spar header
  var cells = targetSheet.getRange(sparHeaderRow, startCol, 1, numCols)
  cells.setBorder(null, null, true, null, null, null)

  targetSheet.autoResizeColumns(startCol, numCols)
}

function printSparHeader(targetSheet, startRow, startCol, numSparrers) {
  targetSheet.getRange(startRow, startCol, 1, 2).setNumberFormat("@")
  targetSheet
    .getRange(startRow, startCol, 1, 2)
    .setValues([["Sparring", "(" + numSparrers + ")"]])
    .setFontSize(16)
    .setFontWeight("bold")
  return 1
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
  physRing,
  numCols
) {
  var cells = targetSheet.getRange(startRow, startCol)
  cells
    .setValue("Ring " + physRing + ' (virtual ring ' + ring + ')')
    .setFontSize(20)
    .setFontWeight("bold")

  // set background color
  var [foregroundcolor, backgroundColor] = getRingBackgroundColors(physRing)
  cells = targetSheet.getRange(startRow, startCol, 1, numCols)
  cells.setBackgroundColor(backgroundColor).setFontColor(foregroundcolor).mergeAcross()

  return 1 // the number of rows printed
}

function printGeneralHeader(sheet, row, col) {
  const headers = [
//    "Order",
    "First",
    "Last",
    "Age",
    "Height",
    "School",
//    "Forms?"
    "Sparring?",
    "gender",
  ]
  cells = sheet.getRange(row, col, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackgroundColor('#f3f3f3')

  return 1
}

// Print out the header cells for one ring
function printFormsHeader(sheet, row, col, numForms, numCols) {
  
  sheet
    .getRange(row, col)
    .setValues("Forms (" + numForms + ")")
    
    .mergeAcross()
    .setFontSize(16)
    .setFontWeight("bold")
    .setNumberFormat("@")

  return 1 // rows in forms header
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
  ]
  var cells = []
  for (let i = 0; i < peopleArr.length; i++) {
    cells.push([])
    for (let j = 0; j < ringHeaders.length; j++) {
      cells[i].push(peopleArr[i][ringHeaders[j]])
    }
  }
  targetSheet
    .getRange(row, col, peopleArr.length, ringHeaders.length)
    .setValues(cells)
    .setHorizontalAlignment("left")

  return peopleArr.length
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
  ]
  for (let i = 0; i < peopleArr.length; i++) {
    for (let j = 0; j < ringHeaders.length; j++) {
      targetSheet
        .getRange(row + i, col + j)
        .setValue(peopleArr[i][ringHeaders[j]])
        .setHorizontalAlignment("left")
    }
  }

  return peopleArr.length
}
