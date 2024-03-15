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

  var startRow = 1

  for (var vRing of Object.keys(virtToPhysMap)) {
    // convert virtual to physical
    var physRingStr = virtToPhysMap[vRing].toString()
    var physArr = physRingStr.match(/\d+|\D+/g)

    // This would be for a horizontal display
    // x = parseInt(physArr[0]) - 1
    // if (physArr[1] == "b") {
    //   y = 1
    // } else if (physArr[1] == "c") {
    //   y = 2
    // } else {
    //   y = 0
    // }
        // x and y are 0-based indices into the table
    // var numColsUsedPerRing = 9
    // var startCol = 1 + numColsUsedPerRing * x
    // var startRow = 1 + 25 * y

    // this is for a vertical display
    x = 0

    var numColsUsedPerRing = 9
    var startCol = 1 + numColsUsedPerRing * x
    

    var peopleInThisVRing = peopleArr.filter((person) => person.vRing == vRing)

    // Maybe we have more rings desclared than people to put in them
    if (peopleInThisVRing.length == 0) {
      continue
    }
    startRow += generateOverviewOneRing(
      targetSheet,
      startCol,
      startRow,
      vRing,
      peopleInThisVRing,
      physRingStr,
      numColsUsedPerRing
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
  numColsUsedPerRing
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
    numColsUsedPerRing
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

  var formRows = generateGenericSubsection(targetSheet, 
    formerArr, 
    "Forms (" + formerArr.length + ")", 
    curRow, 
    startCol, 
    "formOrder",
    numColsUsedPerRing)

  curRow += formRows

  // Get an array of sparrers
  var sparrerArr = []
  var altSparRings = {}
  for (var i = 0; i < peopleArr.length; i++) {
    if (peopleArr[i].sparring.toLowerCase() != "no") {
      sparrerArr.push(peopleArr[i])
      if (peopleArr[i].altSparRing) {
        // If there is an altSparRing, add the key to a hash
        altSparRings[peopleArr[i].altSparRing] = 1
      }
    }
  }

  // turn hash keys into array
  var altSparRingsArr = []
  for (var key in altSparRings) {
    altSparRingsArr.push(key)
  }

  // If hasAltSparRing, then we'll have to do this a couple times
  var sparrerSectionDepth = 0
  
  if (altSparRingsArr.length > 0) {

    for (var altRing = 0; altRing < altSparRingsArr.length; altRing++) {
      var thisSparrerArr = sparrerArr.filter(
        (person) => person.altSparRing == altSparRingsArr[altRing]
      )

      var tmpSparrerSectionDepth = generateGenericSubsection(targetSheet, 
        thisSparrerArr, 
        "Sparring alt ring " + altSparRingsArr[altRing] + " (" + thisSparrerArr.length + ")", 
        curRow, 
        startCol, 
        "sparringOrder",
        numColsUsedPerRing)
        curRow += tmpSparrerSectionDepth
        sparrerSectionDepth += tmpSparrerSectionDepth
    }


  }
  else {
    sparrerArr = sparrerArr.sort(sortBySparringOrder)

    sparrerSectionDepth = generateGenericSubsection(targetSheet, 
      sparrerArr, 
      "Sparring (" + sparrerArr.length + ")", 
      curRow, 
      startCol, 
      "sparringOrder",
      numColsUsedPerRing)
  }
  
  // Border all the way around
  var cells = targetSheet.getRange(
    startRow,
    startCol,
    2 + formerArr.length + sparrerSectionDepth + 1,
    numColsUsedPerRing
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

  return mainHeaderRows + formRows + sparrerSectionDepth

}
  
  
function generateGenericSubsection(
  targetSheet,
  peopleArr,
  headerText,
  startRow,
  startCol,
  orderKey,
  numColsUsedPerRing
) {
  var numPeople = peopleArr.length
  var curRow = startRow
  // ************************
  // Header (ie 'forms')
  targetSheet
  .getRange(curRow, startCol)
  .setValue(headerText)
  .mergeAcross()
  .setFontSize(16)
  .setFontWeight("bold")
  .setNumberFormat("@")
  .setBackgroundColor("#d9d9d9")

  targetSheet
  .getRange(curRow, startCol, 1, numColsUsedPerRing)
  .mergeAcross()
  .setBorder(
    true,
    true,
    true,
    true,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.SOLID
  )

  curRow += 1

  // ********************
  // general header, ie the 'order/first/last' column headings
  var generalHeaderRows = printGeneralHeader(targetSheet, curRow, startCol, orderKey)
  
  targetSheet
  .getRange(curRow, startCol, 1, numColsUsedPerRing)
  .setBorder(
    true,
    true,
    true,
    true,
    null,
    null,
    null,
    SpreadsheetApp.BorderStyle.SOLID
    )
    
  curRow += generalHeaderRows

  // ******************
  // array of people
  peopleArrRows = printPeopleArr(
    targetSheet,
    peopleArr,
    curRow,
    startCol,
    orderKey
  )


  return peopleArrRows + 2
}

function printMainHeader(
  targetSheet,
  startRow,
  startCol,
  ring,
  physRing,
  numColsUsedPerRing
) {
  var cells = targetSheet.getRange(startRow, startCol)
  var cellStr = "Ring " + physRing + ' (virtual ring ' + ring + ')'
  if (globalVariables().displayStyle == "sections") {

    var [physRingNum, sectionLetter, sectionNumber] = splitPhysRing(physRing)

    cellStr = "Ring " + physRingNum + " Section " + sectionNumber + ' (virtual ring ' + ring + ')'
  }
  cells
    .setValue(cellStr)
    .setFontSize(20)
    .setFontWeight("bold")

  // set background color
  var [foregroundcolor, backgroundColor] = getRingBackgroundColors(physRing)
  cells = targetSheet.getRange(startRow, startCol, 1, numColsUsedPerRing)
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
