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
  body.clear().setMarginLeft(18).setMarginBottom(0)

  var paragraph = body.getParagraphs()[0]
  var blob = getImageBlob('logo_orig_dark_letters.png')
  
  var style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
      DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;
  

  var tableSize = {}
  tableSize[DocumentApp.Attribute.FONT_SIZE] = 24
  tableSize[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER

  var row = 0
  var col = 0

  for (var i = 0; i < peopleArr.length; i = i + 1) {

    // if (col==0) {
    //   // put in a new row
    //   buffer.push([])
    // }

    // buffer[row].push(peopleArr[i])

    

    buffer.push(peopleArr[i])
    var pushTable = []

    if ((buffer.length >= numColsPerPage*numRowsPerPage) || (i == peopleArr.length - 1)) {

      // Convert buffer into an array of arrays to put into the table
      
      // for (var tableRow=0; tableRow<numRowsPerPage; tableRow++) {
      //   for (var tableCol=0; tableCol<numColsPerPage; tableCol++) {
      //     if (tableCol + tableRow * numColsPerPage < buffer.length) {
      //       // Here, we have a valid place to put it
      //       pushTable[tableCol][tableRow] = buffer[tableCol + tableRow*numColsPerPage]
      //     }
      //   }
      // }


      var tagTable = body.appendTable()
      //  = body.appendTable([["", ""],["", ""]])
      var lastTagTableRow

      for (j=0; j<buffer.length; j++) {
        if (j%2 == 0) {

          // Add a row, set the row height in points
          lastTagTableRow = tagTable.appendTableRow().setMinimumHeight(2 * 72)
        }

        // Add cells, set the row width in points
        var thisCell = lastTagTableRow.appendTableCell().setWidth(4 * 72)
                    .setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER)

        // The cell was created by default with an existing paragraph. Get that
        // paragraph so there isn't a blank line before the text I want to add.
        var thisParagraph = thisCell.getChild(0)

        // Add in the text. The '\r' is a carriage return. This might have been
        // done another way (maybe append a paragraph) but this works.
        thisParagraph.appendText(buffer[j].sfn + " " + buffer[j].sln + "\r")
        thisParagraph.appendText(buffer[j].school.toString() + "\r")
        var physRing = virtToPhysMap[buffer[j].vRing].toString()
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

      tagTable.setAttributes(tableSize)

      var bottomParagraph = body.appendParagraph("").clear()
      bottomParagraph.appendPageBreak()

      if (i < peopleArr.length - 1) {
        paragraph = body.appendParagraph("").clear()
      }

      buffer = []
    }
  }
  targetDoc.saveAndClose()
}
