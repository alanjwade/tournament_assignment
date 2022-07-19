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

        // Generate one bracket
        generateOneSparringBracket(targetSheet, virtRingPeople, 3, 0, 5)
        // Highlight Semifinal A
  
        // Generate the 3rd place bracket
        generateOneSparringBracket(targetSheet, virtRingPeople, 32, 3, 2)
        // Place header
        generateSparringHeader(targetSheet, 2, 0, sourceSheetName, physRingStr, virtRing)
  
      // Now, virtRingPeople has all the people in one virt ring AND is doing forms
      var body = targetDoc.getBody()
      var style = {}

      throw new Error()
    }
    // Figure out what virtual rings there are.
  
    targetDoc.saveAndClose()
  
    // Get the target doc.
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

function placePeopleInBracket(peopleArr, startRow=0, startCol=0, rounds = 5) {
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

  var position = 0

  //for (var r1 = )

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