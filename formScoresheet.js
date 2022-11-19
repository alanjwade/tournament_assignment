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
  removeImagesFromDoc(targetDoc)

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

  var blob = getImageBlob()

  for (var i=0; i<paragraphsForWatermark.length; i++) {
    paragraphsForWatermark[i].asParagraph().addPositionedImage(blob)
      .setLayout(DocumentApp.PositionedLayout.ABOVE_TEXT)
      .setLeftOffset(0)
      .setTopOffset(0)
      .setWidth(600)
      .setHeight(600)
  }
 
  targetDoc.saveAndClose()
}

// get image blob
function getImageBlob() {
  var image = 'https://drive.google.com/file/d/1HTkFwJWpmlDTLdTRhZZQz4R_NrU0gIj-/view?usp=share_link';
  var fileID = image.match(/[\w\_\-]{25,}/).toString();
  var blob = DriveApp.getFileById(fileID).getBlob();
  return blob
}


// Create a doc with all the forms sheets

function generateFormsSheet(sourceSheetName = "Beginner") {
  var targetDocName = sourceSheetName + " Forms Rings"

  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(sourceSheetName)
  var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet)

  var physToVirtMap = physToVirtMapInv(virtToPhysMap)

  targetDoc = createDocFile(targetDocName)
  var body = targetDoc.getBody()
  body.clear()
  var paragraph = body.getParagraphs()[0]

  // Iterate through the list of sorted physical rings
  for (var physRingStr of sortedPhysRings(virtToPhysMap)) {
    var virtRing = physToVirtMap[physRingStr]

    var virtRingPeople = peopleArr
      .filter(
        (person) =>
          person.vRing == virtRing && person.form.toLowerCase() != "no"
      )
      .sort(sortByFormOrder)

    // Now, virtRingPeople has all the people in one virt ring AND is doing forms
    var timeStamp = createTimeStamp()
    var style = {}
    style[DocumentApp.Attribute.FONT_SIZE] = 8
    var buffer = [
      [
        "First Name",
        "Last Name",
        "School",
        "Virtual Ring",
        "Score 1",
        "Score 2",
        "Score 3",
        "Final Score",
      ],
    ]
    for (var i = 0; i < virtRingPeople.length; i++) {
      buffer.push([
        virtRingPeople[i]["sfn"],
        virtRingPeople[i]["sln"],
        virtRingPeople[i]["school"],
        virtRingPeople[i]["vRing"],
        "",
        "",
        "",
        "",
      ])
    }
    var formTitle =
      sourceSheetName +
      "  Ring " +
      physRingStr
    paragraph.appendText(formTitle)
    paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1)
    paragraph.setSpacingBefore(0)
    formTable = body.appendTable(buffer)
    formTable.setColumnWidth(0, 80)
    formTable.setColumnWidth(1, 80)
    formTable.setColumnWidth(2, 150)
    formTable.setColumnWidth(3, 50)
    formTable.setColumnWidth(4, 50)
    formTable.setColumnWidth(5, 50)
    formTable.setColumnWidth(5, 50)
    formTable.setColumnWidth(5, 50)
    var bottomParagraph = body.appendParagraph(timestamp)

    bottomParagraph.appendPageBreak()
    paragraph = body.appendParagraph("")
  }

  targetDoc.saveAndClose()
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
      "Ring",
      "Score 1",
      "Score 2",
      "Score 3",
      "Final Score",
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
      physRing,
      "",
      "",
      "",
      "",
    ])
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
  formTable.setColumnWidth(1, 90)
  formTable.setColumnWidth(2, 100)
  formTable.setColumnWidth(3, 35)
  formTable.setColumnWidth(4, 50)
  formTable.setColumnWidth(5, 50)
  formTable.setColumnWidth(6, 50)
  formTable.setColumnWidth(7, 50)
  formTable.getRow(0).setAttributes(boldAttr)

  // supposed to center, but the API doesn't do it yet
  formTable.setAttributes(tableStyle)

  var numRows = formTable.getNumRows()
  for (var i=1; i<numRows; i++) {
    var cell = formTable.getCell(i, 7)
    cell.setAttributes(cellStyle)
    // Final cell color
    cell.setBackgroundColor('#efefef')
    // Score cell color
    formTable.getCell(i, 4).setBackgroundColor('#cccccc')
    formTable.getCell(i, 5).setBackgroundColor('#cccccc')
    formTable.getCell(i, 6).setBackgroundColor('#cccccc')
  }

  var bottomParagraph = body.appendParagraph(timeStamp)

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
