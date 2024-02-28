function printNameTagSheet(levelName = "Beginner") {
  var targetDocName = levelName + " Name Tags"
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(levelName)

  var [peopleArr, virtToPhysMap, groupingTable] = readTableIntoArr(sourceSheet)

  peopleArr.sort(sortLastFirst)

  var targetDoc = openOrCreateFileInFolder(
    targetDocName,
    (isSpreadsheet = false),
    (removeFile = false)
  )

  var numRowsPerPage = 4
  var numColsPerPage = 2

  var buffer = []
  var body = targetDoc.getBody()
  body.clear().setMarginLeft(18)

  var paragraph = body.getParagraphs()[0]
  var blob = getImageBlob('logo_orig_dark_letters.png')

  for (var i = 0; i < peopleArr.length; i = i + 1) {

    buffer.push(peopleArr[i])


    if ((buffer.length >= numColsPerPage*numRowsPerPage) || (i == peopleArr.length - 1)) {
      var tagTable = body.appendTable()
      var lastTagTableRow

      for (j=0; j<buffer.length; j++) {
        if (j%2 == 0) {

          // Add a row, set the row height in points
          lastTagTableRow = tagTable.appendTableRow().setMinimumHeight(2 * 72)
        }

        // Add cells, set the row width in points
        var thisCell = lastTagTableRow.appendTableCell().setWidth(4 * 72)
        var thisParagraph = thisCell.appendParagraph("")

        // Fill in the details of one name tag, including logo
        makeNameTagCell(buffer[j], virtToPhysMap, thisParagraph, blob)

      }

      var bottomParagraph = body.appendParagraph("")
      bottomParagraph.appendPageBreak()

      if (i < peopleArr.length - 1) {
        paragraph = body.appendParagraph("")
      }

      buffer = []
    }
  }
  targetDoc.saveAndClose()
}

// Return a paragraph with the right formatting
function makeNameTagCell(person, virtToPhysMap, paragraph, blob) {
  var thisParagraph = paragraph

  thisParagraph.appendText(person.sfn + " ")
  thisParagraph.appendText(person.sln + "\r")
  thisParagraph.appendText(person.school.toString() + "\r")
  var physRing = virtToPhysMap[person.vRing].toString()
  var [fg, bg] = getRingBackgroundColors(physRing)
  thisParagraph.appendText("Ring " + physRing)
               .setForegroundColor(fg).setBackgroundColor(bg)

  thisParagraph.addPositionedImage(blob)
  .setLayout(DocumentApp.PositionedLayout.ABOVE_TEXT)
  .setLeftOffset(2.5*72)
  .setTopOffset(0)
  .setWidth(1.75*72)
  .setHeight(1.5*72)
}