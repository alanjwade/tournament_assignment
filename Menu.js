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
  if (1) {
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
  ui.createMenu("Reorder rings")
    .addItem("reorder rings for all rings", "reorderRingsAll")
    .addSeparator()
    .addItem("reorder rings for beginner rings", "reorderRingsB")
    .addItem("reorder rings for level 1 rings", "reorderRingsL1")
    .addItem("reorder rings for level 2 rings", "reorderRingsL2")
    .addItem("reorder rings for level 3 rings", "reorderRingsL3")
    .addItem("reorder rings for black belt rings", "reorderRingsBB")
    .addToUi()
  ui.createMenu("Gen Collateral")
    .addItem("generate collateral for all rings", "generateCollateralAll")
    .addSeparator()
    .addItem("generate collateral for beginner rings", "generateCollateralB")
    .addItem("generate collateral for level 1 rings", "generateCollateralL1")
    .addItem("generate collateral for level 2 rings", "generateCollateralL2")
    .addItem("generate collateral for level 3 rings", "generateCollateralL3")
    .addItem("generate collateral for black belt rings", "generateCollateralBB")
    .addToUi()
  ui.createMenu("Gen Overview")
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
  ui.createMenu("Gen Checkin Sheets")
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
  ui.createMenu("Gen Score Sheets")
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
  ui.createMenu("Gen Name Tags")
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

function reorderRingsAll() {
  reorderRingsB()
  reorderRingsL1()
  reorderRingsL2()
  reorderRingsL3()
  reorderRingsBB()
}
function reorderRingsB() {
  reorderRings("Beginner")
}
function reorderRingsL1() {
  reorderRings("Level 1")
}
function reorderRingsL2() {
  reorderRings("Level 2")
}
function reorderRingsL3() {
  reorderRings("Level 3")
}
function reorderRingsBB() {
  reorderRings("Black Belt")
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

  // namesToAbbreviations.set("5280 Karate", "5280")
  // namesToAbbreviations.set("Exclusive Martial Arts", "EMA")
  // namesToAbbreviations.set("Personal Achievement Martial Arts", "PAMA")
  // namesToAbbreviations.set("Longmont",     "REMA LM")
  // namesToAbbreviations.set("Broomfield",   "REMA BF")
  // namesToAbbreviations.set("Fort Collins", "REMA FC")
  // namesToAbbreviations.set("Johnstown",    "REMA JT")
  // namesToAbbreviations.set("Success Martial Arts", "SMA")

  namesToAbbreviations.set("exclusive-littleton"      , "EMA LT")
  namesToAbbreviations.set("exclusive-lakewood"       , "EMA LW")
  namesToAbbreviations.set("personal-achievement"     , "PAMA")
  namesToAbbreviations.set("ripple-effect-longmont"   , "REMA LM")
  namesToAbbreviations.set("ripple-effect-broomfield" , "REMA BF")
  namesToAbbreviations.set("ripple-effect-ft-collins" , "REMA FC")
  namesToAbbreviations.set("ripple-effect-johnstown"  , "REMA JT")
  namesToAbbreviations.set("success"                  , "SMA")
  if (namesToAbbreviations.has(schoolName)) {
    abbreviation = namesToAbbreviations.get(schoolName)
  }

  return abbreviation
}

// This is to make the branches of a school look like the same school
function getCommonSchoolAbbreviation(schoolName) {
  
  var abbreviation = schoolName
  
  const namesToAbbreviations = new Map()

  namesToAbbreviations.set("exclusive-littleton"      , "EMA")
  namesToAbbreviations.set("exclusive-lakewood"       , "EMA")
  namesToAbbreviations.set("personal-achievement"     , "PAMA")
  namesToAbbreviations.set("ripple-effect-longmont"   , "REMA")
  namesToAbbreviations.set("ripple-effect-broomfield" , "REMA")
  namesToAbbreviations.set("ripple-effect-ft-collins" , "REMA")
  namesToAbbreviations.set("ripple-effect-johnstown"  , "REMA")
  namesToAbbreviations.set("success"                  , "SMA") 

  if (namesToAbbreviations.has(schoolName)) {
    abbreviation = namesToAbbreviations.get(schoolName)
  }

  return abbreviation
}
// Read table for the purpose of calculating rings.
// Return an array of hashes.
function readTableIntoArr() {
  // Gets sheets data.
  
  
  var peopleSheet = SpreadsheetApp.getActive().getSheetByName("Working")
  var paramSheet = SpreadsheetApp.getActive().getSheetByName("Parameters")
  var parameters = new Map()
  
  // return structure:
  // {peopleSheet, <peopleSheet>}
  // {paramSheet, <paramSheet>}
  // {levelData, {"beginner", {peopleArr, [person0, person1, ...]},
  //                          {virtToPhysMap, <>},
  //                          {maxPeoplePerRingMap, <>}
  //                          {virtToPhysStartRow, <>}
  //                          {maxPeoplePerRingStartRow, <>}
  //                          {virtToPHysStartCol, <>}
  //                          {maxPeoplePerRingStartCol, <>}


  parameters.set("peopleSheet", peopleSheet)
  parameters.set("paramSheet", paramSheet)
  parameters.set("levelData", new Map())
  
  // Get all the people, then sort
  let values = peopleSheet.getDataRange().getValues()

  // Gets the first row of the peopleSheet which is the header row.
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
  let genderCol = headerRowValues.indexOf("Student Gender")
  let vRingCol = headerRowValues.indexOf("Virtual Ring")
  let formOrderCol = headerRowValues.indexOf("Form Order")
  let sparringOrderCol = headerRowValues.indexOf("Sparring Order")
  let divisionCol = headerRowValues.indexOf("Division")
  let physRingCol = headerRowValues.indexOf("PhysRing")

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
  let levels = globalVariables().levels


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
    

    // If divisionCol has one of the levels we're expecting, then add it
    var thisRowDiv = values[i][divisionCol]
    if (levels.includes(thisRowDiv)) {
      if (! parameters.get("levelData").has(thisRowDiv)) {
        parameters.get("levelData").set(thisRowDiv, new Map())
        parameters.get("levelData").get(thisRowDiv).set("peopleArr", new Array())
      }
      parameters.get("levelData").get(thisRowDiv).get("peopleArr").push({
        sfn: values[i][firstNameCol],
        sln: values[i][lastNameCol],
        age: values[i][ageCol],
        grouping: values[i][groupingCol],
        height: values[i][feetCol] + "'" + values[i][inchesCol] + '"',
        heightDec:
        parseInt(values[i][feetCol]) + parseInt(values[i][inchesCol]) / 12.0,
        school: getAbbreviation(values[i][schoolCol]),
        commonSchool: getCommonSchoolAbbreviation(values[i][schoolCol]),
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
        altSparRing: altSparRingVal,
        division: values[i][divisionCol],
        physRingCol: physRingCol,
        physRing: values[i][physRingCol]
      })
    } // values is 0 based
  }
  

  // Get parameters

  let paramValues = paramSheet.getDataRange().getValues()

  for (let i=0; i<paramValues.length; i++) {
    // find the level columns

    // i is a row, j is a column

    for (let j=0; j<paramValues[i].length; j++) {
      if (levels.includes(paramValues[i][j])) {
        var thisLevel = paramValues[i][j]

        var levelMap = parameters.get("levelData").get(thisLevel)

        levelMap.set("virtToPhysStartRow",  i+4)
        levelMap.set("maxPeoplePerRingStartRow", i+4)
        levelMap.set("virtToPhysStartCol", j+1)
        levelMap.set("maxPeoplePerRingStartCol", j+3)

        levelMap.set("virtToPhysMap", new Map())
        levelMap.set("maxPeoplePerRingMap", new Map())

        var levelVRMap = levelMap.get("virtToPhysMap")
        var levelMPMap = levelMap.get("maxPeoplePerRingMap")

        // virt to phys map
        var foundEnd = false

        
        // adjust for going back to the paramValues array
        var row = levelMap.get("virtToPhysStartRow") - 1
        var col = levelMap.get("virtToPhysStartCol") - 1


        while (!foundEnd && (row < paramValues.length)) {
          if (paramValues[row][col] != "") {
            levelVRMap.set(paramValues[row][col], paramValues[row][col+1])
            row += 1
          }
          else {
            foundEnd = true
          }
        }

        // adjust for going back to the paramValues array
        var row = levelMap.get("maxPeoplePerRingStartRow") - 1
        var col = levelMap.get("maxPeoplePerRingStartCol") - 1


        while (!foundEnd && (row < paramValues.length)) {
          if (paramValues[row][col] != "") {
            levelMPMap.set(paramValues[row][col], paramValues[row][col+1])
            row += 1
          }
          else {
            foundEnd = true
          }
        }

        parameters.get("levelData").get(thisLevel).set("virtToPhysMap", levelVRMap)
        parameters.get("levelData").get(thisLevel).set("maxPeoplePerRingMap", levelMPMap)
      }
    }


  }

  // check
  // for (var person of peopleArr) {
  //   if (!(person['vRing'] in virtToPhysMap)) {
  //     console.log("Problem with mapping for " + person['sfn'] + ":" + person['sln'])
  //     var ui = SpreadsheetApp.getUi()
  //     ui.alert('Alert', "Problem with mapping for " + person['sfn'] + ":" + person['sln'],
  //                ui.ButtonSet.OK)
  //   }
  // }


  return parameters
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
  var parameters = readTableIntoArr()
  var peopleArr = parameters.get("levelData").get(level).get("peopleArr")
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
  if (typeof(letter) != "string") {
    return 0
  }
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
  var sectionNumber = convertLetterToNumber(splittedArray[1])

  return [splittedArray[0], splittedArray[1], sectionNumber]
}

function ringDesignator(physRing) {

  var retStr = "Ring " + physRing

  if (globalVariables().displayStyle == "sections") {
    var [physRingNum, groupLetter, groupNumber] = splitPhysRing(physRing)

    retStr = "Ring " + physRingNum + " Group " + groupLetter.toUpperCase()
  }

  return retStr
}

function sortStringsByNumericPrefix(a, b) {
  const aNumberMatch = a.match(/^\d+/);
  const bNumberMatch = b.match(/^\d+/);

  const aNumber = aNumberMatch ? parseInt(aNumberMatch[0]) : 0;
  const bNumber = bNumberMatch ? parseInt(bNumberMatch[0]) : 0;

  if (isNaN(aNumber) || isNaN(bNumber)) {
    // If either number is NaN, sort alphabetically
    return a.localeCompare(b);
  }

  return aNumber - bNumber;
}
