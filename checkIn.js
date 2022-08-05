function printCheckinSheet(levelName = "Beginner") {
  var targetSheetName = levelName + " checkin"
  var targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName)
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(levelName)

  var [peopleArr, virtToPhysMap, groupingTable] = readTableIntoArr(sourceSheet)

  peopleArr.sort(sortLastFirst)

  // 2-d array to store text in before printing it to the sheet
  var buffer = [["First Name", "Last Name", "School", "Physical Ring"]]

  peopleArr.forEach((person) => {
    buffer.push([
      person.sfn,
      person.sln,
      person.school.toString(),
      virtToPhysMap[person.vRing].toString(),
    ])
  })

  targetSheet
    .getRange(1, 1, buffer.length, 4)
    .setValues(buffer)
    .setNumberFormat("@")

  targetDocName = levelName + " docs"

  targetDoc = createDocFile(targetSheetName)
  var body = targetDoc.getBody()
  var style = {}
  style[DocumentApp.Attribute.FONT_SIZE] = 8
  checkinTable = body.clear().appendTable(buffer)
  checkinTable.setAttributes(style)
  checkinTable.setColumnWidth(0, 80)
  checkinTable.setColumnWidth(1, 80)
  checkinTable.setColumnWidth(2, 150)
  checkinTable.setColumnWidth(3, 50)
  style = {}
  style[DocumentApp.Attribute.BOLD] = true
  checkinTable.getRow(0).setAttributes(style)
  targetDoc.saveAndClose()

  console.log("hi")
}

function printFormsSheets(levelName = "Beginner") {
  var targetSheetName = levelName + " forms"
  var targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName)
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(levelName)

  // Get the virtToPhysMap, then invert it. We will do this in order of physical ring.
  var [peopleArr, virtToPhysMap, groupingTable] = readTableIntoArr(sourceSheet)

  var physToVirtMap = physToVirtMapInv(virtToPhysMap)

  // Iterate through the list of sorted physical rings
  for (var physRingstr of sortedPhysRings(virtToPhysMap)) {
    var virtRing = physToVirtMap[physRingstr]

    var virtRingPeople = peopleArr.filter((person) => person.vRing == virtRing)
    console.log("hi")
  }

  peopleArr.sort(sortLastFirst)
}

// Given the virt to phys ring map, return the
// sorted physical rings
function sortedPhysRings(virtToPhysMap) {
  var physToVirtMap = physToVirtMapInv(virtToPhysMap)

  var sortedPhysRingsRet = Object.keys(physToVirtMap).sort(comparePhysRings)
  return sortedPhysRingsRet
}

function sortLastFirst(a, b) {
  var ln_result
  var fn_result
  ln_result = a.sln.localeCompare(b.sln)
  fn_result = a.sfn.localeCompare(b.sfn)

  if (ln_result == 0) {
    return fn_result
  } else {
    return ln_result
  }
}

function createDocFile(fileName) {
  //This query parameter will search for an exact match of the filename with Doc file type
  let params =
    "title='" +
    fileName +
    "' and mimeType = 'application/vnd.google-apps.document'"
  let files = DriveApp.searchFiles(params)
  while (files.hasNext()) {
    //Filename exist
    var file = files.next()
    ///Delete file
    file.setTrashed(true)
  }

  //Create a new file
  let newDoc = DocumentApp.create(fileName)
  return newDoc
}
function createSpreadsheetFile(fileName) {
  //This query parameter will search for an exact match of the filename with Doc file type
  let params =
    "title='" +
    fileName +
    "' and mimeType = 'application/vnd.google-apps.document'"
  let files = DriveApp.searchFiles(params)
  while (files.hasNext()) {
    //Filename exist
    var file = files.next()
    ///Delete file
    file.setTrashed(true)
  }

  //Create a new file
  let newDoc = SpreadsheetApp.create(fileName)
  return newDoc
}

function getSpreadsheetByName(filename) {
  var files = DriveApp.getFilesByName(filename);
  while (files.hasNext()) {
    var file = files.next()
    var ss = SpreadsheetApp.open(file)
    return ss
  }
  return null
}
function getDocByName(filename) {
  var files = DriveApp.getFilesByName(filename);
  while (files.hasNext()) {
    var file = files.next()
    var doc = DocumentApp.openById(file.getId())
    return doc
  }
  return null
}


function createPDFFile(fileName) {
  //This query parameter will search for an exact match of the filename with Doc file type
  
  let files = DriveApp.getFilesByName(fileName)
  while (files.hasNext()) {
    //Filename exist
    var file = files.next()
    ///Delete file
    file.setTrashed(true)
  }

  //Create a new file
  let newDoc = DriveApp.createFile(fileName)
  return newDoc
}

// Returns a list of people in a particular virtual ring
// Optionally also filter on forms or sparring
function getVirtRing(
  peopleArr,
  vRing,
  checkForms = false,
  checkSparring = false
) {
  var retPeopleArr = []

  retPeopleArr = peopleArr.filter((person) => {
    person[vRing] == vRing
  })

  if (checkForms) {
    retPeopleArr = retPeopleArr.filter((person) => {
      person[form]
    })
  }

  if (checkSparring) {
    retPeopleArr = retPeopleArr.filter((person) => {
      person[sparring]
    })
  }
  return retPeopleArr
}

// Invert the virt to phys mapping
function physToVirtMapInv(virtToPhysMap) {
  var retPhysToVirt = {}

  try {
    for (const [virtRing, physRing] of Object.entries(virtToPhysMap)) {
      retPhysToVirt[physRing] = virtRing
    }
  } catch (error) {
    throw new Error("Something's wrong with the virt to phys mapping")
  }
  return retPhysToVirt
}

function comparePhysRings(a, b) {
  aNum = physRingToNumber(a)
  bNum = physRingToNumber(b)

  if (aNum < bNum) {
    return -1
  } else if (aNum == bNum) {
    return 0
  } else {
    return 1
  }
}

function physRingToNumber(physRingStr) {
  var physArr = physRingStr.match(/\d+|\D+/g)
  var ringNumber = parseInt(physArr[0])
  if (physArr.length == 2) {
    if (physArr[1] == "b") {
      ringNumber += 0.5
    } else {
      ringNumber += 0.0
    }
  }
  return ringNumber
}
