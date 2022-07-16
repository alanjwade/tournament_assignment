// Create a doc with all the forms sheets



function generateSparringSheet(sourceSheetName = "Beginner") {
    var targetDocName = sourceSheetName + " Sparring Brackets"
    var targetSheetName = sourceSheetName + " Sparring Brackets"
 
    var sourceSheet = SpreadsheetApp.getActive().getSheetByName(sourceSheetName)
    var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet)
  
    var targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName)

    var physToVirtMap = physToVirtMapInv(virtToPhysMap)
  
    targetDoc = createDocFile(targetDocName)
  
    // Iterate through the list of sorted physical rings
    for (var physRingStr of sortedPhysRings(virtToPhysMap)) {
      var virtRing = physToVirtMap[physRingStr]
  
      var virtRingPeople = peopleArr
        .filter((person) => person.vRing == virtRing && person.sparring != "no")
        .sort(sortBySparringOrder)

        // Generate one bracket
        generateOneSparringBracket(targetSheet, virtRingPeople, 1, 1)
  
      // Now, virtRingPeople has all the people in one virt ring AND is doing forms
      var body = targetDoc.getBody()
      var style = {}
    }
    // Figure out what virtual rings there are.
  
    targetDoc.saveAndClose()
  
    // Get the target doc.
  }


function borderOneCell(sheet, row, col, side) {

  var sides = {
    'top': null,
    'left': null,
    'bottom': null,
    'right': null
  }

  sides[side] = true

  sheet.getRange(row, col).setBorder(sides['top'], sides['left'],
    sides['bottom'], sides['right'], false, false)

  return true
}

function oneFight(sheet, topRow, col, spacing) {
  // Put bottom border on the top one
  borderOneCell(sheet, topRow, col, 'bottom')
  // Put bottom border on the bottom cell
  borderOneCell(sheet, topRow + spacing, col, 'bottom')
  // put right side on the all of them
  sheet.getRange()
}

function getCoordinatesFromRoundPosition(round, position) {
  // Given the round number, and the position within that particular
  // round, return the row, col coordinates.

  // Assume 16 person bracket (call it 5 rounds)

  col = round - 1 // round 1 = col 0
  startRowOffset = Math.pow(2, (round-1)) -1 // round 1: offset 0
                                             // round 2: offset 1
                                             // round 3: offset 3
                                             // round 4: offset 7
                                             // round 5: offset 15

  spacing = round * 2 // round 1: spacing = 2
                      // round 2: spacing = 4 etc

  thisRow = startRowOffset + (position * spacing)

  return [thisRow, col]
}
  
function generateOneSparringBracket(targetSheet, peopleArr, startRow, startCol) {
    // Generate a single sparring bracket

    var totalPeople = peopleArr.length
    // Determine the number of rounds

    //var rounds = Math.ceil(Math.log2(totalPeople))
    var rounds = 5

    var curCol = 0;

    for (var round=1; round<rounds; round++) {
      var maxPositions = Math.pow(2, rounds - round)

      for (var position=0; position < maxPositions; position++) {
        [row, col] = getCoordinatesFromRoundPosition(round, position)

        targetSheet.getRange(row+1, col+1).setBorder(null, null, true, null, null, null)
      }


    }

}