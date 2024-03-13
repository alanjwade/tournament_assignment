// Goal: Create documents that organize participants in various ways.
// Output:
// Physical ring assignments for everyone in a division, both forms and sparring
// Check-in sheet: Alphabetical sheet with all names in a division along with physical ring
// Ring sheet: Ordered list (by criteria included separating people from same schools) of people doing forms
// Sparring bracket: Filled-out bracket of all sparrers in a ring (with 3rd place bracket available and place to fill in 1st through 4th)
// Operations (all per division)
// Assign virtual rings
// Assign physical rings based on map
//   Once the 'calculate' is done, don't run them again. The Ring column is populated and tweaks will be overwritten.
// print rings (prints physical rings of forms and sparring)
// print check in sheet
// print forms sheet ('Ring 5b, Beginner')
// print sparring bracket ('Ring 7a, Black Belt')

// Improvements:
// Can speed up a run by printing things in one setValues call instead of individually
// Don't switch between read and write operations

function onOpen() {
  var ui = SpreadsheetApp.getUi()
  // Or DocumentApp or FormApp.
  if (0) {
    ui.createMenu("Assign Virtual Rings")
      .addItem("Auto Assign all rings", "assignVRingsAll")
      .addSeparator()
      .addItem("Auto Assign beginner rings", "assignVRingsB")
      .addItem("Auto Assign level 1 rings", "assignVRingsL1")
      .addItem("Auto Assign level 2 rings", "assignVRingsL2")
      .addItem("Auto Assign level 3 rings", "assignVRingsL3")
      .addItem("Auto Assign black belt rings", "assignVRingsBB")
      .addToUi()
  }
  ui.createMenu("Generate Collateral")
    .addItem("generate collateral for all rings", "generateCollateralAll")
    .addSeparator()
    .addItem("generate collateral for beginner rings", "generateCollateralB")
    .addItem("generate collateral for level 1 rings", "generateCollateralL1")
    .addItem("generate collateral for level 2 rings", "generateCollateralL2")
    .addItem("generate collateral for level 3 rings", "generateCollateralL3")
    .addItem("generate collateral for black belt rings", "generateCollateralBB")
    .addToUi()
  ui.createMenu("Generate Overview")
    .addItem("generate overview for all rings", "generateOverview")
    .addSeparator()
    .addItem("generate overview for beginner rings", "generateOverviewBRings")
    .addItem("generate overview for level 1 rings", "generateOverviewL1Rings")
    .addItem("generate overview for level 2 rings", "generateOverviewL2Rings")
    .addItem("generate overview for level 3 rings", "generateOverviewL3Rings")
    .addItem(
      "generate overview for black belt rings",
      "generateOverviewBBRings"
    )
    .addSeparator()
    .addToUi()
  ui.createMenu("Generate Checkin Sheets")
    .addItem("generate checkin sheet for all rings", "generateCheckinAll")
    .addSeparator()
    .addItem(
      "generate checkin sheet for beginner rings",
      "generateCheckinBRings"
    )
    .addItem(
      "generate checkin sheet for level 1 rings",
      "generateCheckinL1Rings"
    )
    .addItem(
      "generate checkin sheet for level 2 rings",
      "generateCheckinL2Rings"
    )
    .addItem(
      "generate checkin sheet for level 3 rings",
      "generateCheckinL3Rings"
    )
    .addItem(
      "generate checkin sheet for black belt rings",
      "generateCheckinBBRings"
    )
    .addSeparator()
    .addToUi()
  ui.createMenu("Generate Score Sheets")
    .addItem("generate score sheet for all rings", "generateScoreAll")
    .addSeparator()
    .addItem("generate score sheet for beginner rings", "generateScoreBRings")
    .addItem("generate score sheet for level 1 rings", "generateScoreL1Rings")
    .addItem("generate score sheet for level 2 rings", "generateScoreL2Rings")
    .addItem("generate score sheet for level 3 rings", "generateScoreL3Rings")
    .addItem(
      "generate score sheet for black belt rings",
      "generateScoreBBRings"
    )
    .addSeparator()
    .addToUi()
  ui.createMenu("Generate Name Tags")
    .addItem("generate name tags for all rings", "generateNameTagsAll")
    .addSeparator()
    .addItem("generate name tags for beginner rings", "generateNameTagsBRings")
    .addItem("generate name tags for level 1 rings", "generateNameTagsL1Rings")
    .addItem("generate name tags for level 2 rings", "generateNameTagsL2Rings")
    .addItem("generate name tags for level 3 rings", "generateNameTagsL3Rings")
    .addItem(
      "generate name tags for black belt rings",
      "generateNameTagsBBRings"
    )
    .addToUi()
 
}

function globalVariables() {
  var variables = {
    levels: ["Beginner", "Level 1", "Level 2", "Level 3", "Black Belt"],
    physRingColorMap: {
      1: "#ff0000", // red
      2: "#ffa500", // orange
      3: "#ffff00", // yellow
      4: "#34a853",
      5: "#0000ff",
      6: "#fd2670",
      7: "#8441be", // purpleish
      8: "#999999",
      9: "#000000", // black
      10: "#b68a46",
      11: "#f78db3",
      12: "#6fa8dc",
      13: "#b6d7a8",
      14: "#b4a7d6",
    },
    displayStyle: "sections" // physical rings (1,2,3... or 1a, 1b, 2a, 2b, ...)
                             // sections (ring 1 section 1, ring 1 section 2, etc.)
  }
  return variables
}
function getRingBackgroundColors(physRingStr) {
  // Given a string line 4a, return the foreground and background colors to use.

  var physArr = physRingStr.match(/\d+|\D+/g)
  var physRingNumber = parseInt(physArr[0])
  var backgroundColor = globalVariables().physRingColorMap[physRingNumber]
  var foregroundColor
  if (["#000000", "#0000ff", "#8441be"].includes(backgroundColor)) {
    foregroundColor = "#ffffff" // white
  }
  else {
    foregroundColor = "#000000" // black
  }
  return [foregroundColor, backgroundColor]
}


function generateCollateralAll() {
  generateCollateralB()
  generateCollateralL1()
  generateCollateralL2()
  generateCollateralL3()
  generateCollateralBB()
}
function generateCollateralB() {
  generateCollateral("Beginner")
}
function generateCollateralL1() {
  generateCollateral("Level 1")
}
function generateCollateralL2() {
  generateCollateral("Level 2")
}
function generateCollateralL3() {
  generateCollateral("Level 3")
}
function generateCollateralBB() {
  generateCollateral("Black Belt")
}
function generateCollateral(level) {
  generateOverview(level)
  printCheckinSheet(level)
  printScoresheets(level)
}

function generateCheckinAll() {
  var levels = globalVariables().levels
  levels.forEach((level) => printCheckinSheet(level))
}
function generateCheckinBRings() {
  printCheckinSheet("Beginner")
}
function generateCheckinL1Rings() {
  printCheckinSheet("Level 1")
}
function generateCheckinL2Rings() {
  printCheckinSheet("Level 2")
}
function generateCheckinL3Rings() {
  printCheckinSheet("Level 3")
}
function generateCheckinBBRings() {
  printCheckinSheet("Black Belt")
}

function generateScoreAll() {
  var levels = globalVariables().levels
  levels.forEach((level) => printScoresheets(level))
}
function generateScoreBRings() {
  printScoresheets("Beginner")
}
function generateScoreL1Rings() {
  printScoresheets("Level 1")
}
function generateScoreL2Rings() {
  printScoresheets("Level 2")
}
function generateScoreL3Rings() {
  printScoresheets("Level 3")
}
function generateScoreBBRings() {
  printScoresheets("Black Belt")
}

function createTimeStamp() {
  // Create a text timestamp
  var userTimeZone = CalendarApp.getDefaultCalendar().getTimeZone()
  var thisTimeStr = new Date().toLocaleString("en-US", {
    dateStyle: "long",
    timeStyle: "long",
    timeZone: userTimeZone,
  })
  return "Created " + thisTimeStr
}

function assignVRingsAll() {
  var levels = globalVariables().levels
  levels.forEach((level) => assignVRings(level))
}
function assignVRingsB() {
  assignVRings("Beginner")
}
function assignVRingsL1() {
  assignVRings("Level 1")
}
function assignVRingsL2() {
  assignVRings("Level 2")
}
function assignVRingsL3() {
  assignVRings("Level 3")
}
function assignVRingsBB() {
  assignVRings("Black Belt")
}

function generateOverviewBRings() {
  generateOverview("Beginner")
}
function generateOverviewL1Rings() {
  generateOverview("Level 1")
}
function generateOverviewL2Rings() {
  generateOverview("Level 2")
}
function generateOverviewL3Rings() {
  generateOverview("Level 3")
}
function generateOverviewBBRings() {
  generateOverview("Black Belt")
}

function makeAllRingsCalculated() {
  generateOverview(false, true)
}

function generateNameTagsAll() {
  var levels = globalVariables().levels
  levels.forEach((level) => printNameTagSheet(level))
}
function generateNameTagsBRings() {
  printNameTagSheet("Beginner")
}
function generateNameTagsL1Rings() {
  printNameTagSheet("Level 1")
}
function generateNameTagsL2Rings() {
  printNameTagSheet("Level 2")
}
function generateNameTagsL3Rings() {
  printNameTagSheet("Level 3")
}
function generateNameTagsBBRings() {
  printNameTagSheet("Black Belt")
}

function getId() {
  Browser.msgBox(
    "Spreadsheet key: " + SpreadsheetApp.getActiveSpreadsheet().getId()
  )
}

// Hardcoded abbreviation table, if there is an abbreviation.
function getAbbreviation(schoolName) {
  
  var abbreviation = schoolName
  
  const namesToAbbreviations = new Map()

  namesToAbbreviations.set("5280 Karate Academy", "5280")
  namesToAbbreviations.set("Exclusive Martial Arts", "EMA")
  namesToAbbreviations.set("Personal Achievement Martial Arts", "PAMA")
  namesToAbbreviations.set("Ripple Effect Martial Arts", "REMA")
  namesToAbbreviations.set("Success Martial Arts", "SMA")

  if (namesToAbbreviations.has(schoolName)) {
    abbreviation = namesToAbbreviations.get(schoolName)
  }

  return abbreviation
}

// Read table for the purpose of calculating rings.
// Return an array of hashes.
function readTableIntoArr(sheet) {
  // Gets sheets data.
  let values = sheet.getDataRange().getValues()

  // Gets the first row of the sheet which is the header row.
  let headerRowValues = values[0]
  let feetCol = headerRowValues.indexOf("Feet")
  let firstNameCol = headerRowValues.indexOf("Student First Name")
  let lastNameCol = headerRowValues.indexOf("Student Last Name")
  let ageCol = headerRowValues.indexOf("Student Age")
  let groupingCol = headerRowValues.indexOf("Grouping")
  let inchesCol = headerRowValues.indexOf("Inches")
  let schoolCol = headerRowValues.indexOf("School").toString()
  let formCol = headerRowValues.indexOf("Form")
  let sparringCol = headerRowValues.indexOf("Sparring")
  let genderCol = headerRowValues.indexOf("Gender")
  let vRingCol = headerRowValues.indexOf("Virtual Ring")
  let formOrderCol = headerRowValues.indexOf("Form Order")
  let sparringOrderCol = headerRowValues.indexOf("Sparring Order")

  // If there's an 'Alt Spar Ring' in the headers, then we'll read it.
  // Otherwise, we'll note that there isn't one.

  let altSparRingCol = null
  if (headerRowValues.includes("Alt Spar Ring")) {
    altSparRingCol = headerRowValues.indexOf("Alt Spar Ring")
  }

  // data format:
  //  [{sfn: "jim", sln: "bob", age: 5}, {sfn:"george", sln: "smith", age: 6}, ... ]

  var peopleArr = []
  var endPeopleRow=0

  // start at 1 to avoid header row
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == "") {
      endPeopleRow = i
      break
    }
    // Take care of alt spar ring seperately
    var altSparRingVal
    if (altSparRingCol === null) {
      altSparRingVal = null
    } else {
      altSparRingVal = values[i][altSparRingCol]
      if (altSparRingVal == "") {
        altSparRingVal = null
      }
    }
    peopleArr.push({
      sfn: values[i][firstNameCol],
      sln: values[i][lastNameCol],
      age: values[i][ageCol],
      grouping: values[i][groupingCol],
      height: values[i][feetCol] + "'" + values[i][inchesCol] + '"',
      heightDec:
        parseInt(values[i][feetCol]) + parseInt(values[i][inchesCol]) / 12.0,
      school: getAbbreviation(values[i][schoolCol]),
      form: values[i][formCol],
      sparring: values[i][sparringCol],
      vRing: values[i][vRingCol],
      vRingCol: vRingCol,
      gender: values[i][genderCol],
      originalRow: i + 1,
      formOrder: values[i][formOrderCol],
      sparringOrder: values[i][sparringOrderCol],
      formOrderCol: formOrderCol,
      sparringOrderCol: sparringOrderCol,
      altSparRing: altSparRingVal
    }) // values is 0 based
  }

  // After reading in the people rows, look for the words
  // "maxPeoplePerRing" in a cell in the first column.
  // If found, read the value in the next cell to the right.
  // Put that in maxPeoplePerRing.
  var maxPeoplePerRing = new Map()
  var foundMaxPerRing = false
  var maxPerRingRow

  for (let i=endPeopleRow; i< values.length; i++) {
    if (values[i][0] == "maxPeoplePerRing") {
      maxPeoplePerRing.set('default', values[i][1])
      foundMaxPerRing = true
      maxPerRingRow = i
      break
    }
  }

  if (foundMaxPerRing) {
    for (var i=maxPerRingRow+1; i<values.length; i++) {
      if (values[i][0] == "") {
        break
      }
      else {
        maxPeoplePerRing.set(values[i][0].toString(), values[i][1])
      }
    }
  }

  // Find the grouping table and read it in.
  var groupingTable = []

  var groupingHeaderRow
  var foundHeader = false
  for (
    groupingHeaderRow = 0;
    groupingHeaderRow < values.length;
    groupingHeaderRow++
  ) {
    if (values[groupingHeaderRow][0] == "Grouping Breakdown:") {
      // found the header
      foundHeader = true
      break
    }
  }

  // now read the header if found
  if (foundHeader) {
    for (var mapRow = groupingHeaderRow; mapRow < values.length; mapRow++) {
      if (values[mapRow][0] == "") {
        break
      }
      groupingTable.push([values[mapRow][0], values[mapRow][1]])
    }
  }

  // Find the remapping table and read it in.
  var virtToPhysMap = {}

  var mapHeaderRow
  var foundHeader = false
  for (mapHeaderRow = 0; mapHeaderRow < values.length; mapHeaderRow++) {
    if (values[mapHeaderRow][0] == "Ring Mapping Virtual to Physical") {
      // found the header
      foundHeader = true
      break
    }
  }

  // now read the header if found
  if (foundHeader) {
    for (mapRow = mapHeaderRow + 1; mapRow < values.length; mapRow++) {
      if (values[mapRow][0] == "") {
        break
      }
      virt = values[mapRow][0]
      phys = values[mapRow][1]
      virtToPhysMap[virt] = phys
    }
  }
  return [peopleArr, virtToPhysMap, groupingTable, maxPeoplePerRing, mapHeaderRow]
}

// Get the counts of each scool from an array of people hashes.
function schoolCounts(peopleArr) {
  // Go through each entry, and add increment the count for each school.
  var schoolCountHash = {}

  for (let i = 0; i < peopleArr.length; i++) {
    if (!(peopleArr[i].school in schoolCountHash)) {
      schoolCountHash[peopleArr[i].school] = 1
    } else {
      schoolCountHash[peopleArr[i].school]++
    }
  }
  return schoolCountHash
}

function currentSpreadsheet() {
  //  var thisSpreadsheet = SpreadsheetApp.openById("1PCvAmkn-M8nurOpXY-v7coBzuHbapS0wDumeRpc3E_s");
  var thisSpreadsheet = SpreadsheetApp.getActive()
  return thisSpreadsheet
}

function testSchoolCount() {
  //  var thisSpreadsheet = SpreadsheetApp.openById("1hCB-6ZiJTo0K43WtvoKHx3v4PWmBiUWQqzaC0E6Fx7w");
  var thisSpreadSheet = currentSpreadsheet()
  const catArr = ["Beginner", "Level 1", "Level 2", "Level 3", "Black Belt"]
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName("Beginner")
  var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet)
  var schoolCountHash = schoolCounts(peopleArr)
  var peopleArrSorted = peopleArr.sort(compareByAge)
}

// divide up the groups
// n is the total population
// m is the number of groups
// Returns an array of how many in each group
function divideUpGroups(n, m) {
  var peopleArr = []

  // floor(n/m) + 1 in the first n % m groups
  // floor(n/m) in the remaining rest

  var numFloorPlus1 = n % m
  var numFloor = m - numFloorPlus1

  for (var i = 0; i < numFloorPlus1; i++) {
    peopleArr.push(Math.floor(n / m) + 1)
  }
  for (var j = 0; j < numFloor; j++) {
    peopleArr.push(Math.floor(n / m))
  }
  return peopleArr
}

// Helper function to help sort people by ages.
function compareByAge(a, b) {
  if (a.age < b.age) {
    return -1
  }
  if (a.age > b.age) {
    return 1
  }
  return 0
}

// compare by age rank
function compareByAgeRank(a, b) {
  if (a["age rank"] < b["age rank"]) {
    return -1
  }
  if (a["age rank"] > b["age rank"]) {
    return 1
  }
  return 0
}

function convertLetterToNumber(letter) {
  // Convert the letter to lowercase for case-insensitive conversion
  letter = letter.toLowerCase();

  // Get the character code of the letter
  const charCode = letter.charCodeAt(0);

  // Adjust the character code based on the starting position of 'a' (97)
  // This makes 'a' = 1, 'b' = 2, etc.
  const number = charCode - 96;

  // Check if the input is a valid letter
  if (number < 1 || number > 26) {
    return NaN; // Return NaN for non-letters
  }

  return number;
}

function splitPhysRing(inStr) {

  var regex = new RegExp('([0-9]+)|([a-zA-Z]+)','g')
  var splittedArray = inStr.match(regex)

  return splittedArray
}