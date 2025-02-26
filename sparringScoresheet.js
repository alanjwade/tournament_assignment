

/**
 * Creates a PDF for the customer given sheet.
 * @param {string} ssId - Id of the Google Spreadsheet
 * @param {object} sheet - Sheet to be converted as PDF
 * @param {string} pdfName - File name of the PDF being created
 * @return {file object} PDF file as a blob
 */
function createPDF(ssId, sheet, pdfName) {
  const fr = 0,
    fc = 0,
    lc = 9,
    lr = 27
  const url =
    "https://docs.google.com/spreadsheets/d/" +
    ssId +
    "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" +
    sheet.getSheetId() +
    "&" +
    "r1=" +
    fr +
    "&c1=" +
    fc +
    "&r2=" +
    lr +
    "&c2=" +
    lc

  const params = {
    method: "GET",
    headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() },
  }
  const blob = UrlFetchApp.fetch(url, params)
    .getBlob()
    .setName(pdfName + ".pdf")

  // Gets the folder in Drive where the PDFs are stored.
  const folder = getFolderByName_("CMAA")

  const pdfFile = folder.createFile(blob)
  //   const pdfFile = DriveApp.createFile(blob);

  return pdfFile
}
/**
 * Test function to run getFolderByName_.
 * @prints a Google Drive FolderId.
 */
function test_getFolderByName() {
  // Gets the PDF folder in Drive.
  const folder = getFolderByName_(OUTPUT_FOLDER_NAME)

  console.log(
    `Name: ${folder.getName()}\rID: ${folder.getId()}\rDescription: ${folder.getDescription()}`
  )
  // To automatically delete test folder, uncomment the following code:
  // folder.setTrashed(true);
}
function getFolderByName_(folderName) {
  // Gets the Drive Folder of where the current spreadsheet is located.
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId()
  const parentFolder = DriveApp.getFileById(ssId).getParents().next()

  // Iterates the subfolders to check if the PDF folder already exists.
  const subFolders = parentFolder.getFolders()
  while (subFolders.hasNext()) {
    let folder = subFolders.next()

    // Returns the existing folder if found.
    if (folder.getName() === folderName) {
      return folder
    }
  }
  // Creates a new folder if one does not already exist.
  return parentFolder
    .createFolder(folderName)
    .setDescription(`Created by cmaa application to store PDF output files`)
}

function appendOneSparringScoresheet(
  targetSpreadsheet,
  templateSheet,
  ringPeople,
  virtRing,
  altRingStrAdder,
  physRing,
  level
) {
  // make a new sheet in targetSpreadsheet and populate

  // new sheet
  sheetName = "Ring " + physRing + altRingStrAdder
  if (globalVariables().displayStyle == "sections") {
    var [physRingNum, groupLetter, groupNumber] = splitPhysRing(physRing)
    sheetName = "Ring " + physRingNum + " grp " + groupLetter + altRingStrAdder
  }


  var targetSheet = templateSheet
    .copyTo(targetSpreadsheet)
    .setName(sheetName)

  // assumes the template is already made
  finishOneSparringBracketSheet(
    targetSheet,
    ringPeople,
    0,
    0,
    virtRing,
    physRing,
    level,
    altRingStrAdder
  )

  // Save as pdf

  //  var tempId = tempSpreadsheet.getId()
  //  console.log(tempId)
  //  createPDF(tempId, tempSheet, level + ' bracket')

  //  throw new Error('stop')
  // now tempSheet has the bracket
  //  appendSheetRangeToDocBody(tempSheet, 0, 0, 40, 6, body)
  //  body.appendParagraph("").appendPageBreak()
}

function finishOneSparringBracketSheet(
  targetSheet,
  ringPeople,
  startRow,
  startCol,
  virtRing,
  physRing,
  level,
  altRingStrAdder
) {
  // Put people in a sheet that already had a template applied.
  placePeopleInBracket(targetSheet, ringPeople, startRow + 3, startCol + 0, 5)

  var text = level + " Sparring Bracket " + ringDesignator(physRing) + altRingStrAdder

  // Place header
  generateSparringHeader(
    targetSheet,
    startRow,
    startCol + 0,
    text,
    physRing
  )

  // place timestamp
  targetSheet.getRange(37, 1).setValue(createTimeStamp())
}

function makeOneSparringBracketSheetTemplate(targetSheet, startRow, startCol) {
  // Given a list of people in fighting order, print it to the given sheet.

  // Generate one bracket
  generateOneSparringBracket(targetSheet, startRow + 3, startCol + 0, 5)
  // Highlight Semifinal A and B
  // highlightOneMatch(
  //   targetSheet,
  //   startRow + 3,
  //   startCol + 0,
  //   3,
  //   0,
  //   "#b7e1cd",
  //   "Semifinal Match A"
  // )
  // highlightOneMatch(
  //   targetSheet,
  //   startRow + 3,
  //   startCol + 0,
  //   3,
  //   2,
  //   "#f9cb9c",
  //   "Semifinal Match B"
  // )

  // Highlight semi winners
  var [rowA, colA] = getCoordinatesFromRoundPosition(4, 0)
  var [rowB, colB] = getCoordinatesFromRoundPosition(4, 1)

  // targetSheet
  //   .getRange(startRow + 3 + rowA + 1, startCol + colA + 1)
  //   .setBackground("#b7e1cd")
  // targetSheet
  //   .getRange(startRow + 3 + rowB + 1, startCol + colB + 1)
  //   .setBackground("#f9cb9c")

  // targetSheet
  //   .getRange(startRow + 3 + rowA + 2, startCol + colA + 1)
  //   .setValue("Semifinal Match A Winner")
  //   .setFontWeight('bold')
  //   .setFontSize(11)
  // targetSheet
  //   .getRange(startRow + 3 + rowB + 2, startCol + colB + 1)
  //   .setValue("Semifinal Match B Winner")
  //   .setFontWeight('bold')
  //   .setFontSize(11)

  var [row, col] = getCoordinatesFromRoundPosition(5, 0)
  targetSheet
    .getRange(startRow + 3 + row + 2, startCol + col + 1)
    .setValue("1st")
    .setFontWeight("bold") 
    .setFontSize(14)

  // Generate the 3rd place bracket
  generateOneSparringBracket(targetSheet, startRow + 32, startCol + 3, 2)

  // Add highlights
  var [rowA, colA] = getCoordinatesFromRoundPosition(1, 0)
  var [rowB, colB] = getCoordinatesFromRoundPosition(1, 1)

  // targetSheet
  //   .getRange(startRow + 32 + rowA + 1, startCol + 3 + colA + 1)
  //   .setBackground("#b7e1cd")
  // targetSheet
  //   .getRange(startRow + 32 + rowB + 1, startCol + 3 + colB + 1)
  //   .setBackground("#f9cb9c")

  // targetSheet
  //   .getRange(startRow + 32 + rowA + 2, startCol + 3 + colA + 1)
  //   .setValue("Semifinal Match A Loser")
  //   .setFontWeight('bold')
  //   .setFontSize(11)
  // targetSheet
  //   .getRange(startRow + 32 + rowB + 2, startCol + 3 + colB + 1)
  //   .setValue("Semifinal Match B Loser")
  //   .setFontWeight('bold')
  //   .setFontSize(11)
  var [row, col] = getCoordinatesFromRoundPosition(2, 0)
  targetSheet
    .getRange(startRow + 32 + row + 2, startCol + 3 + col + 1)
    .setValue("3rd")
    .setFontWeight("bold")
    .setFontSize(14)

  // place table
  finalPlaces(targetSheet, 1, 3)

  // Set column widths
  targetSheet.setColumnWidths(1, 5, 200)

  // Set row heights
  targetSheet.setRowHeights(1, 40, 35)

  // Hide gridlines
  targetSheet.setHiddenGridlines(true)

  // insert watermark
  var blob = getImageBlob()
  targetSheet.insertImage(blob, 1, 11, 0, 0)
  .setWidth(900)
  .setHeight(900)
}

function highlightOneMatch(
  targetSheet,
  startRow,
  startCol,
  round,
  startPosition,
  color,
  text, 
  thirdPlaceMatch = false,
  optUpperTxt = '',
  optLowerTxt = ''
) {
  // Highlight one match. round is starting at 1, startPosition is the first fighter.

  var [topRow, col] = getCoordinatesFromRoundPosition(round, startPosition)

  // topRow doesn't include the startRow offset

  var absTopRow = topRow + startRow
  var absCol = col + startCol
  var middleRow

  var topShadeRow
  var numRowsToShade

  if (thirdPlaceMatch) {
    topShadeRow = absTopRow + 23
    numRowsToShade = 2
    middleRow = topShadeRow -1
  }
  else {
    topShadeRow = absTopRow + 1
    numRowsToShade = Math.pow(2, round)
    middleRow = absTopRow + (Math.pow(2, round)) / 2 -1
  }

  targetSheet
    .getRange(topShadeRow, absCol + 1, numRowsToShade, 1)
    .setBackground(color)
    

  targetSheet.getRange(middleRow + 1, absCol + 1).setValue(text)
  .setFontWeight('bold')
  .setFontSize(14)

  if (optUpperTxt != '') {
    targetSheet
    .getRange(topShadeRow - 1, absCol + 1)
    .setValue(optUpperTxt)
    .setFontSize(14)
    .setFontColor('#b7b7b7')
  }
  if (optLowerTxt != '') {
    targetSheet
    .getRange(topShadeRow + numRowsToShade - 1, absCol + 1)
    .setValue(optLowerTxt)
    .setFontSize(14)
    .setFontColor('#b7b7b7')
  }

  return true
}

function generateSparringHeader(
  targetSheet,
  startRow,
  startCol,
  text,
  pRing
) {
  var [foregroundcolor, backgroundColor] = getRingBackgroundColors(pRing)
  targetSheet.getRange(startRow + 1, startCol + 1, 1, 5).mergeAcross()

  // Set the header and the colors
  targetSheet
    .getRange(startRow + 1, startCol + 1)
    .setValue(text)
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackgroundColor(backgroundColor)
    .setFontColor(foregroundcolor)
}

function borderOneCell(sheet, row, col, side) {
  var sides = {
    top: null,
    left: null,
    bottom: null,
    right: null,
  }

  sides[side] = true

  sheet
    .getRange(row, col)
    .setBorder(
      sides["top"],
      sides["left"],
      sides["bottom"],
      sides["right"],
      false,
      false
    )

  return true
}

function oneFight(sheet, topRow, col, spacing) {
  // Put bottom border on the top one
  borderOneCell(sheet, topRow, col, "bottom")
  // Put bottom border on the bottom cell
  borderOneCell(sheet, topRow + spacing, col, "bottom")
  // put right side on the all of them
  sheet.getRange()
}

function getCoordinatesFromRoundPosition(round, position) {
  // Given the round number, and the position within that particular
  // round, return the row, col coordinates.

  // Assume 16 person bracket (call it 5 rounds)

  col = round - 1 // round 1 = col 0
  startRowOffset = Math.pow(2, round - 1) - 1 // round 1: offset 0
  // round 2: offset 1
  // round 3: offset 3
  // round 4: offset 7
  // round 5: offset 15

  spacing = Math.pow(2, round) // round 1: spacing = 2
  // round 2: spacing = 4
  // round 3: spacing = 8

  thisRow = startRowOffset + position * spacing

  return [thisRow, col]
}

function placePeopleInBracket(
  targetSheet,
  peopleArr,
  startRow = 0,
  startCol = 0,
  rounds = 5
) {
  var totalPeople = peopleArr.length

  // Find which round we're going to start in

  var startRound = 0
  var start2Round = 0

  // This is the number playing in the 'play-in' round,
  // or the first round if there are a power of 2 number of people.
  var numStartRound = 0
  // This is the number playing in the next round.
  var numStart2Round = 0

  var slotsInNextRound

  for (var round = 1; round <= rounds; round++) {
    var slotsInStartRound = Math.pow(2, rounds - round)
    slotsInNextRound = Math.pow(2, rounds - (round + 1))
    if (totalPeople == slotsInStartRound) {
      startRound = round
      start2Round = round
      numStartRound = totalPeople
      numStart2Round = 0
      break
    } else if ((totalPeople < slotsInStartRound) & (totalPeople > slotsInNextRound)) {
      startRound = round
      start2Round = round + 1

      // need an even number to fight in the start round
      numStartRound = 2 * (totalPeople - slotsInNextRound)
      numStart2Round = totalPeople - numStartRound

      break
    }
  }

  // we put numStartRound in round 'startRound' and numStart2Round in 'start2Round'
  for (var personIndex = 0; personIndex < totalPeople; personIndex++) {
    var thisRound
    var thisPosition
    if (personIndex < numStart2Round) {
      thisRound = start2Round
      thisPosition = personIndex
    } else {
      thisRound = startRound
      thisPosition = (personIndex - numStart2Round) + (slotsInStartRound - numStartRound)
    }

    var [row, col] = getCoordinatesFromRoundPosition(thisRound, thisPosition)

    targetSheet
      .getRange(startRow + row + 1, startCol + col + 1)
      .setValue(
        peopleArr[personIndex]["sfn"] + " " + peopleArr[personIndex]["sln"]
        + " (" + peopleArr[personIndex]["school"] + ")"
      )

    // set background color for match
//  
  }

  // Set the background colors and add 'Match #x' to each match
  var matchIdx = 1
  for (var round=startRound; round<rounds; round++) {
    var slotsInThisRound = Math.pow(2, rounds - round)
    var startPosition
    if (round == startRound) {
      startPosition = slotsInThisRound - numStartRound
    }
    else {
      startPosition = 0
    }

    for (var position = startPosition; position <slotsInThisRound; position += 2) {
      var thisMatch
      var semiTopTxt
      var semiBotTxt
      if (round == rounds - 1) {
        if (matchIdx == 1) {
          // final round, but there were only 2 competitors
          thisMatch = 1
          semiTopTxt = ''
          semiBotTxt = ''
        }
        else if (matchIdx == 2) {
          // final round, but there were only 3 competitors
          thisMatch = 2
          semiTopTxt = ''
          semiBotTxt = 'Winner Match #1'
        }
        else {
           // final round, 4 or more competitors
           thisMatch = matchIdx + 1
           semiTopTxt = 'Winner Match #' + (matchIdx - 2)
           semiBotTxt = 'Winner Match #' + (matchIdx - 1)
        }
      }
      else {
        thisMatch = matchIdx++
        semiTopTxt = ''
        semiBotTxt = ''
       }
      var sizeOfBackground = Math.pow(2, thisRound)
      highlightOneMatch(
        targetSheet,
        startRow + 1,
        startCol + 0,
        round,
        position,
        getMatchBackgroundColor(round, position/2),
        'Match #' + thisMatch,
        thirdPlaceMatch = false,
        semiTopTxt,
        semiBotTxt
      )
    }
  }

  // Highlight the 3rd place match
  if (matchIdx > 2) {
    highlightOneMatch(
      targetSheet,
      startRow + 1,
      startCol + 0,
      rounds - 1,
      0,
      getMatchBackgroundColor(rounds - 1, 0),
      'Match #' + matchIdx, //should be one before the final
      thirdPlaceMatch = true,
      optUpperTxt = 'Loser Match #' + (matchIdx - 2),
      optLowerTxt = 'Loser Match #' + (matchIdx - 1)
    )
  }
}

function generateOneSparringBracket(
  targetSheet,
  startRow = 0,
  startCol = 0,
  rounds = 5
) {
  // Generate a single sparring bracket

  var curCol = 0

  for (var round = 1; round <= rounds; round++) {
    var maxPositions = Math.pow(2, rounds - round)

    var lastRow = 0
    for (var position = 0; position < maxPositions; position++) {
      ;[row, col] = getCoordinatesFromRoundPosition(round, position)

      targetSheet
        .getRange(startRow + row + 1, startCol + col + 1)
        .setBorder(
          null,
          null,
          true,
          null,
          null,
          null,
          "black",
          SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        )

      // Make the right borders
      // If we're an odd position, make a right border between this and the last position
      if (position % 2 != 0) {
        for (var borderRows = lastRow + 1; borderRows <= row; borderRows++) {
          targetSheet
            .getRange(startRow + borderRows + 1, startCol + col + 1)
            .setBorder(
              null,
              null,
              null,
              true,
              null,
              null,
              "black",
              SpreadsheetApp.BorderStyle.SOLID_THICK
            )
        }
      }

      lastRow = row
    }
  }
}

function finalPlaces(targetSheet, startRow, startCol) {
  var buffer = []
  buffer.push(["Final Places", ""])
  buffer.push(["1st", ""])
  buffer.push(["2nd", ""])
  buffer.push(["3rd", ""])

  targetSheet.getRange(startRow + 1, startCol + 1, 4, 2).setValues(buffer)
  targetSheet
    .getRange(startRow + 2, startCol + 1, 3, 2)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "black",
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    )
  targetSheet.getRange(startRow + 1, startCol + 1, 1, 1)
    .setFontSize(18)
    .setFontWeight('bold')
}

function getMatchBackgroundColor(round, boutZB) {
  // Return a hex color for the round (1-5) and the match position number
  // This should be consistent for all sizes

  var colorMap = [
    ['#e6f7ff', '#cceeff', '#b3e6ff', '#99ddff', '#80d4ff', '#66ccff', '#4dc3ff', '#33bbff'],
    ['#e6fff2', '#ccffe6', '#b3ffd9', '#99ffcc'],
    ['#fff5e6', '#ffcc80'],
    ['#feeaea']
  ]

  return colorMap[round-1][boutZB]

}