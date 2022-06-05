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
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Assign Virtual Rings')
      .addItem('Auto Assign all rings', 'assignVRingsAll')
      .addSeparator()
      .addItem('Auto Assign beginner rings', 'assignVRingsB')
      .addItem('Auto Assign level 1 rings', 'assignVRingsL1')
      .addItem('Auto Assign level 2 rings', 'assignVRingsL2')
      .addItem('Auto Assign level 3 rings', 'assignVRingsL3')
      .addItem('Auto Assign black belt rings', 'assignVRingsBB')
      .addToUi();
  ui.createMenu('Print Rings')
      .addItem('populate all rings', 'makeRings')
      .addItem('populate beginner rings', 'makeBRings')
      .addItem('populate level 1 rings', 'makeL1Rings')
      .addItem('populate level 2 rings', 'makeL2Rings')
      .addItem('populate level 3 rings', 'makeL3Rings')
      .addItem('populate black belt rings', 'makeBBRings')
      .addSeparator()
      .addToUi();
}

function globalVariables() {
  var variables = {
    levels: ['Beginner', 'Level 1', 'Level 2', 'Level 3', 'Black Belt'],
    phyRingColorMap: {1: "red",
                      2: "orange",
                      3: "yellow",
                      4: "#34a853",
                      5: "#0000ff",
                      6: "#fd2670",
                      7: "#8441be",
                      8: "#999999",
                      9: "black",
                      10: "#b68a46",
                      11: "#f78db3",
                      12: "#6fa8dc",
                      13: "#b6d7a8",
                      14: "#b4a7d6"}
  }
  return variables
}

function assignVRingsAll() {
  var levels = globalVariables().levels
  levels.forEach(level => assignVRings(level))
}
function assignVRingsB() {
  assignVRings('Beginner')
}
function assignVRingsL1() {
  assignVRings('Level 1')
}
function assignVRingsL2() {
  assignVRings('Level 2')
}
function assignVRingsL3() {
  assignVRings('Level 3')
}
function assignVRingsBB() {
  assignVRings('Black Belt')
}

function printRemappedAllRings() {
    makeRings(false, false, useRemapping = true);
}
function printRemappedBRings() {
    makeRings('Beginner', false, useRemapping = true);
}
function printRemappedL1Rings() {
    makeRings('Level 1', false, useRemapping = true);
}
function printRemappedL2Rings() {
    makeRings('Level 2', false, useRemapping = true);
}
function printRemappedL3Rings() {
    makeRings('Level 3', false, useRemapping = true);
}
function printRemappedBBRings() {
    makeRings('Black Belt', false, useRemapping = true);
}

function makeBRings() {
    makeRings('Beginner');
}
function makeL1Rings() {
    makeRings('Level 1');
}
function makeL2Rings() {
    makeRings('Level 2');
}
function makeL3Rings() {
    makeRings('Level 3');
}
function makeBBRings() {
    makeRings('Black Belt');
}

function makeAllRingsCalculated() {
  makeRings(false, true);
}

function makeBRingsCalculated() {
  makeRings('Beginner', true);
}
function makeL1Calculated() {
  makeRings('Level 1', true);
}
function makeL2Calculated() {
  makeRings('Level 2', true);
}
function makeL3Calculated() {
  makeRings('Level 3', true);
}
function makeBBRingsCalculated() {
  makeRings('Black Belt', true);
}

function genAgeSchoolBeginner() {
  assignVRingAgeSchool("Beginner");
}
function genAgeSchoolL1() {
  assignVRingAgeSchool("Level 1");
}
function genAgeSchoolL2() {
  assignVRingAgeSchool("Level 2");
}
function genAgeSchoolL3() {
  assignVRingAgeSchool("Level 3");
}
function genAgeSchoolBB() {
  assignVRingAgeSchool("Black Belt");
}

function genAgeSchoolAll() {
  assignVRingAgeSchool("Beginner");
  assignVRingAgeSchool("Level 1");
  assignVRingAgeSchool("Level 2");
  assignVRingAgeSchool("Level 3");
  assignVRingAgeSchool("Black Belt");
}

function ajwgetname() {
  var sheet = SpreadsheetApp.getActive().getName();
  console.log(sheet);
  return sheet;
}
function getId() {
  Browser.msgBox('Spreadsheet key: ' + SpreadsheetApp.getActiveSpreadsheet().getId());
}



// Read table for the purpose of calculating rings.
// Return an array of hashes.
function readTableIntoArr(sheet) {
  // Gets sheets data.
  let values = sheet.getDataRange().getValues();

  // Gets the first row of the sheet which is the header row.
  let headerRowValues = values[0];
  let feetCol = headerRowValues.indexOf("Feet");
  let firstNameCol = headerRowValues.indexOf("Student First Name");
  let lastNameCol = headerRowValues.indexOf("Student Last Name");
  let ageCol = headerRowValues.indexOf("Student Age");
  let groupingCol = headerRowValues.indexOf("Grouping")
  let inchesCol = headerRowValues.indexOf("Inches");
  let schoolCol = headerRowValues.indexOf("School").toString();
  let formCol = headerRowValues.indexOf("Form");
  let sparringCol = headerRowValues.indexOf("Sparring");
  let genderCol = headerRowValues.indexOf("Gender");
  let vRingCol = headerRowValues.indexOf("Virtual Ring");
  let formOrderCol = headerRowValues.indexOf("Form Order")
  let sparringOrderCol = headerRowValues.indexOf("Sparring Order")

  // data format:
  //  [{sfn: "jim", sln: "bob", age: 5}, {sfn:"george", sln: "smith", age: 6}, ... ]

  var peopleArr = [];

  // start at 1 to avoid header row
  for (let i=1; i<values.length; i++) {
    if (values[i][0] == '') {
      break;
    }
    peopleArr.push({sfn: values[i][firstNameCol],
                        sln: values[i][lastNameCol],
                        age: values[i][ageCol],
                        grouping: values[i][groupingCol],
                        height: values[i][feetCol] + "\'" + values[i][inchesCol] + "\"",
                        heightDec: parseInt(values[i][feetCol]) + (parseInt(values[i][inchesCol]) / 12.0),
                        school: values[i][schoolCol],
                        form: values[i][formCol],
                        sparring: values[i][sparringCol],
                        vRing: values[i][vRingCol],
                        vRingCol: vRingCol,
                        gender: values[i][genderCol],
                        originalRow: i+1,
                        formOrder: values[i][formOrderCol],
                        sparringOrder: values[i][sparringOrderCol],
                        formOrderCol: formOrderCol,
                        sparringOrderCol: sparringOrderCol}) // values is 0 based

  }

  // Find the grouping table and read it in.
  var groupingTable = [];

  var groupingHeaderRow;
  var foundHeader = false;
  for (groupingHeaderRow = 0; groupingHeaderRow < values.length; groupingHeaderRow++) {
    if (values[groupingHeaderRow][0] == "Grouping Breakdown:") {
      // found the header
      foundHeader = true;
      break;
    }
  }

  // now read the header if found
  if (foundHeader) {
    for (var mapRow = groupingHeaderRow; mapRow < values.length; mapRow++) {
      if (values[mapRow][0] == "") {
        break;
      }
      groupingTable.push([values[mapRow][0], values[mapRow][1]])
    }
  }

  // Find the remapping table and read it in.
  var virtToPhysMap = {};

  var mapHeaderRow;
  var foundHeader = false;
  for (mapHeaderRow = 0; mapHeaderRow < values.length; mapHeaderRow++) {
    if (values[mapHeaderRow][0] == "Ring Mapping Virtual to Physical") {
      // found the header
      foundHeader = true;
      break;
    }
  }

  // now read the header if found
  if (foundHeader) {
    for (mapRow = mapHeaderRow + 1; mapRow < values.length; mapRow++) {
      if (values[mapRow][0] == "") {
        break;
      }
      virt = values[mapRow][0];
      phys = values[mapRow][1];
      virtToPhysMap[virt] = phys;
    }
  }
  return [peopleArr, virtToPhysMap, groupingTable]              
}

// Get the counts of each scool from an array of people hashes.
function schoolCounts(peopleArr) {

  // Go through each entry, and add increment the count for each school.
  var schoolCountHash = {};

  for (let i=0; i<peopleArr.length; i++) {

    if (!(peopleArr[i].school in schoolCountHash)) {
      schoolCountHash[peopleArr[i].school] = 1;
    }
    else {
      schoolCountHash[peopleArr[i].school]++;
    }
  }
  return schoolCountHash;
}


function currentSpreadsheet() {
//  var thisSpreadsheet = SpreadsheetApp.openById("1PCvAmkn-M8nurOpXY-v7coBzuHbapS0wDumeRpc3E_s");
  var thisSpreadsheet = SpreadsheetApp.getActive();
  return thisSpreadsheet;
}

function testSchoolCount() {
//  var thisSpreadsheet = SpreadsheetApp.openById("1hCB-6ZiJTo0K43WtvoKHx3v4PWmBiUWQqzaC0E6Fx7w");
  var thisSpreadSheet = currentSpreadsheet();
  const catArr = ['Beginner', 'Level 1', 'Level 2', 'Level 3', 'Black Belt'];
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName("Beginner");
  var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet);
  var schoolCountHash = schoolCounts(peopleArr);
  var peopleArrSorted = peopleArr.sort(compareByAge);
}





// divide up the groups
// n is the total population
// m is the number of groups
// Returns an array of how many in each group
function divideUpGroups(n, m) {
  var peopleArr = [];

  // floor(n/m) + 1 in the first n % m groups
  // floor(n/m) in the remaining rest

  var numFloorPlus1 = n % m;
  var numFloor = (m - numFloorPlus1);

  for (var i=0; i<numFloorPlus1; i++) {
    peopleArr.push(Math.floor(n/m) + 1);
  }
  for (var j=0; j<numFloor; j++) {
    peopleArr.push(Math.floor(n/m))
  }
  return peopleArr;
}




// Helper function to help sort people by ages.
function compareByAge(a, b) {
  if ( a.age < b.age ) {
    return -1;
  }
  if ( a.age > b.age ) {
    return 1;
  }
  return 0;
}

// compare by age rank
function compareByAgeRank(a, b) {
  if ( a["age rank"] < b["age rank"] ) {
    return -1;
  }
  if ( a["age rank"] > b["age rank"] ) {
    return 1;
  }
  return 0;
}

