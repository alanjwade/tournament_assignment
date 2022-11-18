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
  targetSheet.clearFormats()

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
    } else if (physArr[1] == "c") {
      y = 2
    } else {
      y = 0
    }
    // x and y are 0-based indices into the table
    var numCols = 9
    var startCol = 1 + numCols * x
    var startRow = 1 + 25 * y

    var peopleInThisVRing = peopleArr.filter((person) => person.vRing == vRing)

    // Maybe we have more rings desclared than people to put in them
    if (peopleInThisVRing.length == 0) {
      continue
    }
    generateOverviewOneRing(
      targetSheet,
      startCol,
      startRow,
      vRing,
      peopleInThisVRing,
      physRingStr,
      numCols
    )

    // Generate a timestamp
    targetSheet.getRange(53, 1).setValue(createTimeStamp())
    targetSheet.getRange(53, 1, 1, 5).mergeAcross()
  }
}

function generateOverviewOneRing(
  targetSheet,
  startCol,
  startRow,
  ringId,
  peopleArr,
  physRing,
  numCols
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
    numCols
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

  curRow += printGeneralHeader(targetSheet, curRow, startCol)


  // *********************************
  // Array of form people
  // *********************************
  var formRows = printPeopleArr(targetSheet, formerArr, curRow, startCol, "formOrder")

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
    numSparrers,
    numCols
  )

  var sparHeaderRow = curRow
  curRow += sparHeaderRows
  curRow += printGeneralHeader(targetSheet, curRow, startCol)

  // *******************************
  // Sparring array
  // *******************************
  var numSparrers = printPeopleArr(
    targetSheet,
    sparrerArr,
    curRow,
    startCol,
    "sparringOrder"
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

  // Before the 'Forms' or 'Sparring' header
  targetSheet.getRange(formHeaderRow, startCol, 1, numCols)
  .setBorder(
    true,
    null,
    null,
    null,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.SOLID_THICK
  )
  targetSheet.getRange(sparHeaderRow, startCol, 1, numCols)
  .setBorder(
    true,
    null,
    null,
    null,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.SOLID_THICK
  )
  
  // After the 'Name' row
  targetSheet.getRange(formHeaderRow + 1, startCol, 1, numCols)
  .setBorder(
    null,
    null,
    true,
    null,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.SOLID
  )
  targetSheet.getRange(sparHeaderRow + 1, startCol, 1, numCols)
  .setBorder(
    null,
    null,
    true,
    null,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.SOLID
  )
 
  targetSheet.setColumnWidth(startCol, 50)
  targetSheet.autoResizeColumns(startCol+1, numCols-1)
}


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
    "Order",
    "First",
    "Last",
    "Age",
    "Height",
    "School",
    "Forms?",
    "Sparring?",
    "gender",
  ]
  cells = sheet.getRange(row, col, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackgroundColor('#f3f3f3')

  return 1
}

// Print out the header cells for one ring
function printFormsHeader(sheet, row, col, numForms, mergeCols) {
  
  sheet
    .getRange(row, col)
    .setValue("Forms (" + numForms + ")")
    .mergeAcross()
    .setFontSize(16)
    .setFontWeight("bold")
    .setNumberFormat("@")
    .setBackgroundColor("#d9d9d9")

  sheet.getRange(row, col, 1, mergeCols).mergeAcross()

  return 1 // rows in forms header
}

function printSparHeader(sheet, row, col, numSparrers, mergeCols) {
  sheet
    .getRange(row, col)
    .setValue("Sparring (" + numSparrers + ")")
    .mergeAcross()
    .setFontSize(16)
    .setFontWeight("bold")
    .setNumberFormat("@")
    .setBackgroundColor("#d9d9d9")

    sheet.getRange(row, col, 1, mergeCols).mergeAcross()
    return 1

}

// Fill out the info for one ring. Just the data, not the headers.
function printPeopleArr(targetSheet, peopleArr, row, col, orderKey) {
  const ringHeaders = [
    orderKey,
    "sfn",
    "sln",
    "age",
    "height",
    "school",
    "form",
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
  if (peopleArr.length > 0) {
    targetSheet
      .getRange(row, col, peopleArr.length, ringHeaders.length)
      .setValues(cells)
      .setHorizontalAlignment("left")
  }
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
