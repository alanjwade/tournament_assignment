function printNameTagSheet(level = "Beginner") {
  var targetDocName = level + " Name Tags"

  //var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet)
  var parameters = readTableIntoArr()

  var virtToPhysMap = parameters.get("levelData").get(level).get("virtToPhysMap")
  var peopleArr = parameters.get("levelData").get(level).get("peopleArr")

  peopleArr.sort(sortLastFirst)

  var targetDoc = openOrCreateFileInFolder(
    targetDocName,
    (isSpreadsheet = false),
    (removeFile = false)
  )

  var numRowsPerPage = 4
  var numColsPerPage = 2

  const pointsPerInch = 72

  var labelInfo = {}
  labelInfo["heightIn"] = 2+ 1/3 + 0.25
  labelInfo["widthIn"] = 3.375 + 3/8
  labelInfo["leftMarginIn"] = (8.5 - 2*3.375 - (3/8)) / 2
  labelInfo["topMarginIn"] = (11 - 4 * labelInfo["heightIn"]) / 2

  var heightPoints = labelInfo["heightIn"] * pointsPerInch
  var widthPoints = labelInfo["widthIn"] * pointsPerInch
  var leftMarginPoints = labelInfo["leftMarginIn"] * pointsPerInch
  var topMarginPoints = labelInfo["topMarginIn"] * pointsPerInch
  
  tinyStyle = {}
  tinyStyle[DocumentApp.Attribute.FONT_SIZE] = 1

  var buffer = []
  var body = targetDoc.getBody()
  body.clear().setMarginLeft(leftMarginPoints)
              .setMarginBottom(0)
              .setMarginTop(topMarginPoints)
              .setAttributes(tinyStyle)

//   footer = targetDoc.addFooter()
//  footer.appendParagraph(createTimeStamp())

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

      var tmpParagraphs = body.getParagraphs()
      var lastParagraph = tmpParagraphs[tmpParagraphs.length - 1]
      lastParagraph.setAttributes(tinyStyle)
      var lastParagraphIndex = body.getChildIndex(lastParagraph)


      var tagTable = body.insertTable(lastParagraphIndex)
//                         .setBorderColor("#c0bfbc")
                           .setBorderColor("#ffffff")
//  = body.appendTable([["", ""],["", ""]])
      var lastTagTableRow

      for (var j=0; j<buffer.length; j++) {
        if (j%2 == 0) {

          // Add a row, set the row height in points
          lastTagTableRow = tagTable.appendTableRow().setMinimumHeight(heightPoints)
        }

        // Add cells, set the row width in points
        var thisCell = lastTagTableRow.appendTableCell().setWidth(widthPoints)
                    .setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER)

        // The cell was created by default with an existing paragraph. Get that
        // paragraph so there isn't a blank line before the text I want to add.
        var thisParagraph = thisCell.getChild(0)

        // Add in the text. The '\r' is a carriage return. This might have been
        // done another way (maybe append a paragraph) but this works.
        var physRing = virtToPhysMap.get(buffer[j].vRing).toString()
        var [fg, bg] = getRingBackgroundColors(physRing)
        thisParagraph.appendText(buffer[j].sfn + " " + buffer[j].sln + "\r")
                     .appendText(buffer[j].school.toString() + "\r")
                     .appendText(level + "\r")
        thisParagraph.appendText(ringDesignator(physRing))
                     .setForegroundColor(fg).setBackgroundColor(bg)
      
        thisParagraph.addPositionedImage(blob)
        .setLayout(DocumentApp.PositionedLayout.ABOVE_TEXT)
        .setLeftOffset(.65 * widthPoints)
        .setTopOffset(heightPoints - .50 * widthPoints) // guess at vert position
        .setWidth(.25 * widthPoints)
        .setHeight(.25 * widthPoints)

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
