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
      .filter((person) => person.vRing == virtRing && person.form != "no")
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
