function printScoresheets(level = "Beginner") {
  // Print the combined forms/sparring scoresheet for particular level
  //
  var targetDocName = level + " Forms Score Sheets"

  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(level)

  //var targetDoc = getDocByName(targetDocName)
  var targetDoc = openOrCreateFileInFolder(
    targetDocName,
    (isSpreadsheet = false)
  )
  targetDoc.getBody().clear()
    // This needs the Google Docs API, which might have to be reenabled
  //  removeImagesFromDoc(targetDoc)

  var footer = targetDoc.getFooter()
  if (footer) {
    footer.removeFromParent()
  }
  footer = targetDoc.addFooter()
  
  var centerStyle = {}
  centerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.RIGHT;

  //footer.appendImage(getImageBlob('logo.png')).setAttributes(centerStyle)
  footer.appendParagraph(createTimeStamp())

  var targetBody = targetDoc.getBody()
  var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet)

  var physToVirtMap = physToVirtMapInv(virtToPhysMap)

  // Iterate through the list of sorted physical rings
  // For now, use a new spreadsheet.
  var targetSheetName = level + " Sparring Score Sheets"
  //var targetSpreadsheet = createSpreadsheetFile(targetSheetName)
  //var targetSpreadsheet = getSpreadsheetByName(targetSheetName)
  var targetSpreadsheet = openOrCreateFileInFolder(
    targetSheetName,
    (isSpreadsheet = true)
  )

  if (targetSpreadsheet == null) {
    throw new Error()
  }

  // Add one dummy one here
  var tmp = targetSpreadsheet.getSheetByName("dummy")
  if (tmp == null) {
    targetSpreadsheet.insertSheet("dummy")
  }

  var tmp = targetSpreadsheet.getSheetByName("template")
  if (tmp != null) {
    targetSpreadsheet.deleteSheet(tmp)
  }

  // delete all the sheets to start over, except one called 'template'
  var allSheets = targetSpreadsheet.getSheets()

  templateSheet = targetSpreadsheet.insertSheet("template")
  makeOneSparringBracketSheetTemplate(templateSheet, 0, 0)

  for (var i = 0; i < allSheets.length; i++) {
    targetSpreadsheet.deleteSheet(allSheets[i])
  }

  var paragraphsForWatermark = []
  for (var physRingStr of sortedPhysRings(virtToPhysMap)) {
    var virtRing = physToVirtMap[physRingStr]

    // Get all the people in one virtRing, whether forms, sparring, or both
    var virtRingPeople = peopleArr.filter((person) => person.vRing == virtRing)

    // If there's no one in this virtual ring, skip
    if (virtRingPeople.length == 0) {
      continue
    }

    // Make the forms scoresheet
    // filter on doing forms and then sort
    var formsPeople = virtRingPeople
      .filter((person) => person.form.toLowerCase() != "no")
      .sort(sortByFormOrder)

    paragraphsForWatermark.push(appendOneFormsScoresheet(
                                targetBody,
                                formsPeople,
                                virtRing,
                                physRingStr,
                                level
                                )
                               )
    console.log("Finished with forms ring " + physRingStr)

    // Make the sparring scoresheet
    // filter on doing forms and then sort
    var sparringPeople = virtRingPeople
      .filter((person) => person.sparring.toLowerCase() != "no")
      .sort(sortBySparringOrder)

    appendOneSparringScoresheet(
      targetSpreadsheet,
      templateSheet,
      sparringPeople,
      virtRing,
      physRingStr,
      level
    )

    console.log("Finished with sparring ring " + physRingStr)
  }
  // Remove the dummy sheet. Can't do that before if it's the only one,
  // so wait until the others are made.
  targetSpreadsheet.deleteSheet(templateSheet)

  var paragraphs = targetDoc.getBody().getParagraphs()

  console.log('Num paragraphs: ' + paragraphs.length)


 //This will put the watermark on every page
  var blob = getImageBlob()
  for (var i=0; i<paragraphsForWatermark.length; i++) {
    paragraphsForWatermark[i].asParagraph().addPositionedImage(blob)
      .setLayout(DocumentApp.PositionedLayout.ABOVE_TEXT)
      .setLeftOffset(0)
      .setTopOffset(50)
      .setWidth(700)
      .setHeight(700)
  }
 
  targetDoc.saveAndClose()
}

// Get the image file 'watermark.png' in this directory
function getImageBlob(filename = 'watermark.png') {
  // Get this spreadsheet
  var ss = SpreadsheetApp.getActive()

  // Get the folder. Hopefully there's just one. Pick it
  var parentFolder = DriveApp.getFileById(ss.getId()).getParents().next()
  console.log("looking for " + filename + " in " + parentFolder.getName())

  // See if there's the 'filename' in this directory. It will be the ID if it's there.
  var file = fileExistsInFolder(filename, parentFolder)

  var blob = DriveApp.getFileById(file.getId()).getBlob()
  return blob
   
}


function appendOneFormsScoresheet(body, ringPeople, virtRing, physRing, level) {
  // Given a Body object, append one ring worth of forms scoresheet

  var style = {}
  style[DocumentApp.Attribute.FONT_SIZE] = 8
  var buffer = [
    [
      "First Name",
      "Last Name",
      "School",
      "Score 1",
      "Score 2",
      "Score 3",
      "Final Score",
      "Final Place"
    ],
  ]

  body.setMarginLeft(40)
  body.setMarginRight(25)
  var boldAttr = {}
  boldAttr[DocumentApp.Attribute.BOLD] = true
  var unboldAttr = {}
  unboldAttr[DocumentApp.Attribute.BOLD] = false
  var tableStyle = {};
  tableStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.BORDER_WIDTH] = 2.25; 
  cellStyle[DocumentApp.Attribute.BORDER_COLOR] = '#ffffff';

  // Get the last paragraph, so we don't end up with a space before the first form
  var paragraphs = body.getParagraphs()
  var topParagraph = paragraphs[paragraphs.length - 1]
  var timeStamp = createTimeStamp()
  for (var i = 0; i < ringPeople.length; i++) {
    buffer.push([
      ringPeople[i]["sfn"],
      ringPeople[i]["sln"],
      ringPeople[i]["school"],
      "",
      "",
      "",
      "",
      "",
    ])
  }
  // Add in a few extra
  for (var i=0; i<0; i++) {
    buffer.push(["","","","","","","",""])
  }
  var formTitle =
    level +  " Ring " + physRing
  var [foregroundcolor, backgroundColor] = getRingBackgroundColors(physRing)
  
  var titleText = topParagraph.appendText(formTitle)
  topParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  titleText.setBackgroundColor(backgroundColor)
  titleText.setForegroundColor(foregroundcolor)
 
  topParagraph.setSpacingBefore(0)
  dummyParagraph = body.appendParagraph("")
  formTable = body.appendTable(buffer)
  formTable.setAttributes(unboldAttr)
  formTable.setColumnWidth(0, 80)
  formTable.setColumnWidth(1, 100)
  formTable.setColumnWidth(2, 90)
  formTable.setColumnWidth(3, 50)
  formTable.setColumnWidth(4, 50)
  formTable.setColumnWidth(5, 50)
  formTable.setColumnWidth(6, 50)
  formTable.setColumnWidth(7, 50)
  formTable.getRow(0).setAttributes(boldAttr)

  placeBuffer = ([['Final Place', 'Name'],
                  ['1', ''],
                  ['2', ''],
                  ['3', '']])
  placeTable = body.appendTable(placeBuffer)
  placeTable.setColumnWidth(0, 70)
  placeTable.setColumnWidth(1, 200)
  placeTable.getRow(0).setAttributes(boldAttr)

  // supposed to center, but the API doesn't do it yet
  formTable.setAttributes(tableStyle)

  var numRows = formTable.getNumRows()
  for (var i=1; i<numRows; i++) {
    // Score cell color
    formTable.getCell(i, 3).setBackgroundColor('#cccccc')
    formTable.getCell(i, 4).setBackgroundColor('#cccccc')
    formTable.getCell(i, 5).setBackgroundColor('#cccccc')
    formTable.getCell(i, 6).setBackgroundColor('#efefef')
    
    var row = formTable.getRow(i)
    for (var c = 0; c < row.getNumCells(); c++) {
      row.getCell(c).setPaddingTop(2).setPaddingBottom(0);
    }}

  var bottomParagraph = body.appendParagraph("")

  bottomParagraph.appendPageBreak()
  //paragraph = body.appendParagraph("")

  return dummyParagraph
}

// sort function for form order.
function sortByFormOrder(a, b) {
  return a.formOrder - b.formOrder
}

// sort function for form order.
function sortBySparringOrder(a, b) {
  return a.sparringOrder - b.sparringOrder
}
