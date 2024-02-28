function printNameTagSheet(levelName = "Beginner") {
  var targetDocName = levelName + " Name Tags";
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(levelName);

  var [peopleArr, virtToPhysMap, groupingTable] = readTableIntoArr(sourceSheet);

  peopleArr.sort(sortLastFirst);

  var targetDoc = openOrCreateFileInFolder(
    targetDocName,
    (isSpreadsheet = false),
    (removeFile = false)
  );

  var numRowsPerPage = 4;
  var cols = 2;

  var buffer = [];
  var body = targetDoc.getBody();
  body.clear();

  var paragraph = body.getParagraphs()[0];

  var row = 0
  var col = 0;

  for (var i = 0; i < peopleArr.length; i = i + 2) {
    // Do 2 at a time and special case an odd one at the end

    // Each cell is one name tag
    thisRow = [
      peopleArr[i].sfn +
        " " +
        peopleArr[i].sln +
        "\n" +
        peopleArr[i].school.toString() +
        "\n" +
        virtToPhysMap[peopleArr[i].vRing].toString(),
    ];

    if (i + 1 < peopleArr.length) {
      thisRow.push(
        peopleArr[i+1].sfn +
          " " +
          peopleArr[i+1].sln +
          "\n" +
          peopleArr[i+1].school.toString() +
          "\n" +
          "ring " + 
          virtToPhysMap[peopleArr[i+1].vRing].toString()
      );
    }

    buffer.push(thisRow);
    row = row + 1;

    if (row > numRowsPerPage) {
      tagTable = body.appendTable(buffer);
      var bottomParagraph = body.appendParagraph("");
      bottomParagraph.appendPageBreak();

      if (i < peopleArr.length - 1) {
        paragraph = body.appendParagraph("");
      }

      buffer = [];
      row = 0;
    }
  }
  targetDoc.saveAndClose();
}
