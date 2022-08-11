function printScoresheets(level = "Beginner") {
  // Print the combined forms/sparring scoresheet for particular level
  // 
  var targetDocName = level + " Forms Score Sheets";

  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(level);

  //var targetDoc = getDocByName(targetDocName)
  var targetDoc = openOrCreateFileInFolder(targetDocName, isSpreadsheet = false);
  targetDoc.getBody().clear();

  var targetBody = targetDoc.getBody();
  var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet);

  var physToVirtMap = physToVirtMapInv(virtToPhysMap);

  // Iterate through the list of sorted physical rings
  // For now, use a new spreadsheet.
  var targetSheetName = level + " Sparring Score Sheets";
  //var targetSpreadsheet = createSpreadsheetFile(targetSheetName)
  //var targetSpreadsheet = getSpreadsheetByName(targetSheetName)
  var targetSpreadsheet = openOrCreateFileInFolder(targetSheetName, isSpreadsheet = true);

  if (targetSpreadsheet == null) {
    throw new Error;
  }

  // Add one dummy one here
  var tmp = targetSpreadsheet.getSheetByName('template');
  if (tmp != null) {
    targetSpreadsheet.deleteSheet(tmp);
  }

  // delete all the sheets to start over, except one called 'template'
  var allSheets = targetSpreadsheet.getSheets();

  templateSheet = targetSpreadsheet.insertSheet('template');
  makeOneSparringBracketSheetTemplate(templateSheet, 0, 0);

  for (var i = 0; i < allSheets.length; i++) {
    targetSpreadsheet.deleteSheet(allSheets[i]);
  }



  for (var physRingStr of sortedPhysRings(virtToPhysMap)) {
    var virtRing = physToVirtMap[physRingStr];


    // Get all the people in one virtRing, whether forms, sparring, or both
    var virtRingPeople = peopleArr
      .filter((person) => person.vRing == virtRing);

    // Make the forms scoresheet
    // filter on doing forms and then sort
    var formsPeople = virtRingPeople.filter((person) => person.form.toLowerCase() != "no")
      .sort(sortByFormOrder);

    appendOneFormsScoresheet(targetBody, formsPeople, virtRing, physRingStr, level);
    console.log('Finished with forms ring ' + physRingStr);

    // Make the sparring scoresheet
    // filter on doing forms and then sort
    var sparringPeople = virtRingPeople.filter((person) => person.sparring.toLowerCase() != "no")
      .sort(sortBySparringOrder);

    appendOneSparringScoresheet(targetSpreadsheet, templateSheet, sparringPeople, virtRing, physRingStr, level);

    console.log('Finished with sparring ring ' + physRingStr);
  }
  // Remove the dummy sheet. Can't do that before if it's the only one,
  // so wait until the others are made.
  targetSpreadsheet.deleteSheet(templateSheet);


  targetDoc.saveAndClose();


}


// Create a doc with all the forms sheets


function generateFormsSheet(sourceSheetName = "Beginner") {
  var targetDocName = sourceSheetName + " Forms Rings"

  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(sourceSheetName)
  var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet)

  var physToVirtMap = physToVirtMapInv(virtToPhysMap)

  targetDoc = createDocFile(targetDocName)

  // Iterate through the list of sorted physical rings
  for (var physRingStr of sortedPhysRings(virtToPhysMap)) {
    var virtRing = physToVirtMap[physRingStr]

    var virtRingPeople = peopleArr
      .filter((person) => person.vRing == virtRing && person.form.toLowerCase() != "no")
      .sort(sortByFormOrder)

    // Now, virtRingPeople has all the people in one virt ring AND is doing forms
    var body = targetDoc.getBody()
    var style = {}
    style[DocumentApp.Attribute.FONT_SIZE] = 8
    var buffer = [
      ["First Name", "Last Name", "School", "Virtual Ring", "Score 1", "Score 2", "Score 3", "Final Score"],
    ]
    for (var i = 0; i < virtRingPeople.length; i++) {
      buffer.push([
        virtRingPeople[i]["sfn"],
        virtRingPeople[i]["sln"],
        virtRingPeople[i]["school"],
        virtRingPeople[i]["vRing"],
        "", "", "", ""
      ])
    }
    var formTitle =
      sourceSheetName +
      " Virt Ring " +
      virtRing +
      " Physical Ring " +
      physRingStr
    body.appendParagraph(formTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1)
    body.appendTable(buffer)
    body.appendParagraph("").appendPageBreak()
  }
  // Figure out what virtual rings there are.

  targetDoc.saveAndClose()

  // Get the target doc.
}

function appendOneFormsScoresheet(body, ringPeople, virtRing, physRing, level) {
  // Given a Body object, append one ring worth of forms scoresheet

  var style = {}
  style[DocumentApp.Attribute.FONT_SIZE] = 8
  var buffer = [
    ["First Name", "Last Name", "School", "Ring", "Score 1", "Score 2", "Score 3", "Final Score"]]

  for (var i = 0; i < ringPeople.length; i++) {
    buffer.push([
      ringPeople[i]["sfn"],
      ringPeople[i]["sln"],
      ringPeople[i]["school"],
      physRing,
      "", "", "", ""
    ])
  }
  var formTitle =
    level +
    " Virtual Ring " +
    virtRing +
    " Physical Ring " +
    physRing
  body.appendParagraph(formTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1)
  body.appendTable(buffer)
  body.appendParagraph("").appendPageBreak()
}

// sort function for form order.
function sortByFormOrder(a, b) {
  return a.formOrder - b.formOrder
}

// sort function for form order.
function sortBySparringOrder(a, b) {
  return a.sparringOrder - b.sparringOrder
}
