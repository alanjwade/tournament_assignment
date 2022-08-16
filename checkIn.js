function printCheckinSheet(levelName = "Beginner") {
  var targetDocName = levelName + " Checkin"
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(levelName)

  var [peopleArr, virtToPhysMap, groupingTable] = readTableIntoArr(sourceSheet)

  peopleArr.sort(sortLastFirst)

  // 2-d array to store text in before printing it to the sheet
  var headBuffer = ["First Name", "Last Name", "School", "Ring"]

  //  targetDoc = createDocFile(targetSheetName)
  var targetDoc = openOrCreateFileInFolder(
    targetDocName,
    (isSpreadsheet = false)
  )
  var tableSize = {}
  tableSize[DocumentApp.Attribute.FONT_SIZE] = 12
  var headerSize = {}
  headerSize[DocumentApp.Attribute.FONT_SIZE] = 14
  var checkTitle = levelName + " Checkin Sheet"

  var boldAttr = {}
  boldAttr[DocumentApp.Attribute.BOLD] = true
  var unboldAttr = {}
  unboldAttr[DocumentApp.Attribute.BOLD] = false

  // Put 25 people on a page
  const numPeoplePerPage = 25
  const totalPages = Math.ceil(peopleArr.length / numPeoplePerPage)
  var curPage = 1
  var timeStamp = createTimeStamp()
  var buffer = []
  // initial paragraph
  var body = targetDoc.getBody()
  body.clear()
  var paragraph = body.getParagraphs()[0]
  for (var i = 0; i < peopleArr.length; i++) {
    // After 25 or the end, put in a new page
    buffer.push([
      peopleArr[i].sfn,
      peopleArr[i].sln,
      peopleArr[i].school.toString(),
      virtToPhysMap[peopleArr[i].vRing].toString(),
    ])
    if (
      i % numPeoplePerPage == numPeoplePerPage - 1 ||
      i == peopleArr.length - 1
    ) {
      buffer.unshift(headBuffer)

      //body.appendParagraph(checkTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1)
      paragraph.appendText(checkTitle + " Page " + curPage++ + "/" + totalPages)
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1)
      paragraph.setSpacingBefore(0)
      checkinTable = body.appendTable(buffer)
      checkinTable.setAttributes(unboldAttr)
      checkinTable.setAttributes(tableSize)
      checkinTable.setColumnWidth(0, 80)
      checkinTable.setColumnWidth(1, 100)
      checkinTable.setColumnWidth(2, 150)
      checkinTable.setColumnWidth(3, 50)
      checkinTable.getRow(0).setAttributes(boldAttr)
      checkinTable.getRow(0).setAttributes(headerSize)

      // set the padding to 0 all around for all the cells
      for (var r=0; r<checkinTable.getNumRows(); r++) {
        for (var c=0; c<4; c++) {
          checkinTable.getCell(r,c).setPaddingTop(0).setPaddingBottom(0)
        }
      }

      var bottomParagraph = body.appendParagraph(timeStamp)
      bottomParagraph.appendPageBreak()

      if (i < peopleArr.length - 1) {
        paragraph = body.appendParagraph("")
      }
      buffer = []
    }
  }
  targetDoc.saveAndClose()
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

function fileExistsInFolder(filename, folder) {
  // from   https://stackoverflow.com/questions/39685232/google-script-test-for-file-existance

  console.log("looking in " + folder.getName() + " for " + filename)
  var file = folder.getFilesByName(filename)
  console.log("hasNext: " + file.hasNext())
  if (file.hasNext()) {
    return file.next()
  } else {
    return false
  }
}

function openOrCreateFileInFolder(filename, isSpreadsheet) {
  // Get this spreadsheet
  var ss = SpreadsheetApp.getActive()

  // Get the folder. Hopefully there's just one. Pick it
  var parentFolder = DriveApp.getFileById(ss.getId()).getParents().next()
  console.log("looking for " + filename + " in " + parentFolder.getName())

  // See if there's the 'filename' in this directory.
  var file = fileExistsInFolder(filename, parentFolder)
  if (file) {
    console.log("Found " + filename + ", returning it")
    if (isSpreadsheet) {
      return SpreadsheetApp.open(file)
    } else {
      return DocumentApp.openById(file.getId())
    }
  } else {
    // Create the file
    console.log("Did not find " + filename)
    if (isSpreadsheet) {
      var newDoc = SpreadsheetApp.create(filename)
    } else {
      var newDoc = DocumentApp.create(filename)
    }

    // Move it to the folder
    var newFile = DriveApp.getFileById(newDoc.getId())

    newFile.moveTo(parentFolder)

    return newDoc
  }
}

function getSpreadsheetByName(filename) {
  var files = DriveApp.getFilesByName(filename)
  while (files.hasNext()) {
    var file = files.next()
    var ss = SpreadsheetApp.open(file)
    return ss
  }
  return null
}
function getDocByName(filename) {
  var files = DriveApp.getFilesByName(filename)
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
