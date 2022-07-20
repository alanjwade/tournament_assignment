// Create a doc with all the forms sheets



function generateSparringSheet(sourceSheetName = "Beginner") {
    var targetDocName = sourceSheetName + " Sparring Brackets"
    var targetSheetName = sourceSheetName + " Sparring Brackets"
 
    var sourceSheet = SpreadsheetApp.getActive().getSheetByName(sourceSheetName)
    var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet)
  
    var targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName)
    targetSheet.clear()

    var physToVirtMap = physToVirtMapInv(virtToPhysMap)
  
    targetDoc = createDocFile(targetDocName)
  
    // Iterate through the list of sorted physical rings
    for (var physRingStr of sortedPhysRings(virtToPhysMap)) {
      var virtRing = physToVirtMap[physRingStr]
  
      var virtRingPeople = peopleArr
        .filter((person) => person.vRing == virtRing && person.sparring != "no")
        .sort(sortBySparringOrder)

        generateOneSparringBracketSheet(targetSheet, virtRingPeople, 0, 0, physRingStr, virtRing, sourceSheetName)
  
      // Now, virtRingPeople has all the people in one virt ring AND is doing forms
      var body = targetDoc.getBody()
      var style = {}

      throw new Error()
    }
    // Figure out what virtual rings there are.
  
    targetDoc.saveAndClose()
  
    // Get the target doc.
}

function generateOneSparringBracketSheet(targetSheet, virtRingPeople, startRow, startCol, physRingStr, virtRing, level) {
    // Given a list of people in fighting order, print it to the given sheet.

    // Generate one bracket
    generateOneSparringBracket(targetSheet, virtRingPeople, startRow + 3, startCol + 0, 5)
    placePeopleInBracket(targetSheet, virtRingPeople, startRow + 3, startCol + 0, 5)
    // Highlight Semifinal A and B
    highlightOneMatch(targetSheet, startRow + 3, startCol + 0, 3, 0, "#b7e1cd", 'Semifinal Match A')
    highlightOneMatch(targetSheet, startRow + 3, startCol + 0, 3, 2, "#f9cb9c", 'Semifinal Match B')
    
    // Highlight semi winners
    var [rowA, colA] = getCoordinatesFromRoundPosition(4, 0)
    var [rowB, colB] = getCoordinatesFromRoundPosition(4, 1)

    targetSheet.getRange(startRow + 3 + rowA + 1, startCol + colA + 1).setBackground("#b7e1cd")
    targetSheet.getRange(startRow + 3 + rowB + 1, startCol + colB + 1).setBackground("#f9cb9c")

    targetSheet.getRange(startRow + 3 + rowA + 2, startCol + colA + 1).setValue("Semifinal Match A Winner")
    targetSheet.getRange(startRow + 3 + rowB + 2, startCol + colB + 1).setValue("Semifinal Match B Winner")

    var [row, col] = getCoordinatesFromRoundPosition(5, 0)
    targetSheet.getRange(startRow + 3 + row + 2, startCol + col + 1).setValue("1st")



    // Generate the 3rd place bracket
    generateOneSparringBracket(targetSheet, virtRingPeople, startRow + 32, startCol + 3, 2)

    // Add highlights
    var [rowA, colA] = getCoordinatesFromRoundPosition(1, 0)
    var [rowB, colB] = getCoordinatesFromRoundPosition(1, 1)

    targetSheet.getRange(startRow + 32 + rowA + 1, startCol + 3 + colA + 1).setBackground("#b7e1cd")
    targetSheet.getRange(startRow + 32 + rowB + 1, startCol + 3 + colB + 1).setBackground("#f9cb9c")

    targetSheet.getRange(startRow + 32 + rowA + 2, startCol + 3 + colA + 1).setValue("Semifinal Match A Loser")
    targetSheet.getRange(startRow + 32 + rowB + 2, startCol + 3 + colB + 1).setValue("Semifinal Match B Loser")
    var [row, col] = getCoordinatesFromRoundPosition(2, 0)
    targetSheet.getRange(startRow + 32 + row + 2, startCol + 3 + col + 1).setValue("3rd")

    // Place header
    generateSparringHeader(targetSheet, startRow + 2, startCol + 0, level, physRingStr, virtRing)
}

function highlightOneMatch (targetSheet, startRow, startCol, round, startPosition, color, text) {
  // Highlight one match. round is starting at 1, startPosition is the first fighter.

  var [topRow, col] = getCoordinatesFromRoundPosition(round, startPosition)

  // topRow doesn't include the startRow offset

  var absTopRow = topRow + startRow
  var absCol = col + startCol

  targetSheet.getRange(absTopRow + 1, absCol + 1, Math.pow(2, round)+1, 1).setBackground(color)

  var middleRow = absTopRow + (Math.pow(2, round) / 2)

  targetSheet.getRange(middleRow + 1, absCol + 1).setValue(text)

  return true

}

function generateSparringHeader(targetSheet, startRow, startCol, text, pRing, vRing) {
  targetSheet.getRange(startRow + 1, startCol + 1).setValue(text + ' Sparring Bracket Ring ' + pRing)
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

  spacing = Math.pow(2, round) // round 1: spacing = 2
                               // round 2: spacing = 4
                               // round 3: spacing = 8

  thisRow = startRowOffset + (position * spacing)

  return [thisRow, col]
}

function placePeopleInBracket(targetSheet, peopleArr, startRow=0, startCol=0, rounds = 5) {
  var totalPeople = peopleArr.length

  // Find which round we're going to start in

  var startRound = 0
  var start2Round = 0

  // This is the number playing in the 'play-in' round,
  // or the first round if there are a power of 2 number of people.
  var numStartRound = 0
  // This is the number playing in the next round.
  var numStart2Round = 0

  for (var round = 1; round <= rounds; round++) {
    var numThisRound = Math.pow(2, rounds-round)
    var numNextRound = Math.pow(2, rounds - (round + 1))
    if (totalPeople == numThisRound) {
      startRound = round
      start2Round = round
      numStartRound = totalPeople
      numStart2Round = 0
      break
    }
    else if ((totalPeople < numThisRound) & (totalPeople > numNextRound)) {
      startRound = round
      start2Round = round + 1

      // need an even number to fight in the start round
      numStartRound = 2 * (totalPeople - numNextRound)
      numStart2Round = totalPeople - numStartRound
      break
    }
  }

  // we put numStartRound in round 'startRound' and numStart2Round in 'start2Round'

  for (var personIndex = 0; personIndex < totalPeople; personIndex++) {
    var thisRound
    var thisPosition
    if (personIndex < numStartRound) {
      thisRound = startRound
      thisPosition = personIndex
    }
    else {
      thisRound = start2Round
      thisPosition = personIndex - numStart2Round
    }



    var [row, col] = getCoordinatesFromRoundPosition(thisRound, thisPosition)

    targetSheet.getRange(startRow + row+1, startCol + col+1).setValue(peopleArr[personIndex]['sln'])

  }

}
  
function generateOneSparringBracket(targetSheet, peopleArr, startRow = 0, startCol = 0, rounds = 5) {
    // Generate a single sparring bracket

    var totalPeople = peopleArr.length
    // Determine the number of rounds

    //var rounds = Math.ceil(Math.log2(totalPeople))

    var curCol = 0;

    for (var round=1; round<=rounds; round++) {
      var maxPositions = Math.pow(2, rounds - round)

      var lastRow = 0
      for (var position=0; position < maxPositions; position++) {
        [row, col] = getCoordinatesFromRoundPosition(round, position)

        targetSheet.getRange(startRow + row + 1, startCol + col + 1).setBorder(null, null, true, null, null, null)

        // Make the right borders
        // If we're an odd position, make a right border between this and the last position
        if (position % 2 != 0) {
          for (var borderRows = lastRow+1; borderRows <= row; borderRows++) {
            targetSheet.getRange(startRow + borderRows + 1, startCol + col + 1).setBorder(null, null, null, true, null, null)
          }
        }

        lastRow = row
      }
    }
}