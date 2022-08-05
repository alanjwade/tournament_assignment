function assignVRings(sourceSheetName = "Beginner") {
  var targetSheetName = sourceSheetName + " forms"
  var targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName)
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(sourceSheetName)
  var [peopleArr, virtToPhysMap] = readTableIntoArr(sourceSheet)

  var separatedByGroup = separateIntoGroups(peopleArr)

  var groupingsSortedByAgeRank = {}

  for (grouping in separatedByGroup) {
    groupingsSortedByAgeRank[grouping] = rankOneGroup(
      separatedByGroup[grouping]
    )
  }

  // Now groupingsSortedByAgeRank is all the groupings separated out in a hash,
  // with the list sorted by age ranks.
  // Now, divide each grouping up into the number of rings they will need.

  var startRing = 1
  var vRingHash = {} // hash of (vring, list of people)
  for (var grouping in groupingsSortedByAgeRank) {
    // choose the number of rings to use
    var numRingsThisGroup = Math.ceil(
      groupingsSortedByAgeRank[grouping].length / 10
    ) // max 10 per ring

    var vRingHashTmp = divideOneGroupingIntoVRings(
      groupingsSortedByAgeRank[grouping],
      startRing,
      numRingsThisGroup
    )

    for (const [key, value] of Object.entries(vRingHashTmp)) {
      vRingHash[key] = value
    }

    startRing = startRing + numRingsThisGroup
  }

  // Finally, populate the spreadsheet
  for (const [vRing, vRingPeopleArr] of Object.entries(vRingHash)) {
    for (var i = 0; i < vRingPeopleArr.length; i++) {
      var row = vRingPeopleArr[i].originalRow
      var col = vRingPeopleArr[i].vRingCol + 1
      sourceSheet.getRange(row, col).setValue(vRing)
    }
  }

  // Figure out and assign the order for forms.
  for (const [vRing, vRingPeopleArr] of Object.entries(vRingHash)) {
    var formPeople = vRingPeopleArr.filter((person) => person.form != "no")
    var inFormOrder = applySortOrder(formPeople, sortByNameHashcode, "formRank")
    for (var index = 0; index < inFormOrder.length; index++) {
      sourceSheet
        .getRange(
          inFormOrder[index]["originalRow"],
          inFormOrder[index]["formOrderCol"] + 1
        )
        .setValue(index + 1)
    }
  }

  // Figure out and assign the order for sparring.
  for (const [vRing, vRingPeopleArr] of Object.entries(vRingHash)) {
    var sparPeople = vRingPeopleArr.filter((person) => person.sparring != "no")
    var inSparOrder = applySortOrder(sparPeople, sortByHeight, "sparRank")
    for (var index = 0; index < inSparOrder.length; index++) {
      sourceSheet
        .getRange(
          inSparOrder[index]["originalRow"],
          inSparOrder[index]["sparringOrderCol"] + 1
        )
        .setValue(index + 1)
    }
  }

  return vRingHash
}

// Calculate the ring based on grouping and school.
// Put the results back in the original page in calculated ring col.
function assignVRingAgeSchool(sourceSheetName) {
  var sourceSheet = SpreadsheetApp.getActive().getSheetByName(sourceSheetName)
  var peopleArr = readTableIntoArr(sourceSheet)

  // Put this in a table like this:
  // sortedAgeBySchool = {"Success": [p1, p2, p3...],
  //                      "PA"     : [p1, p2, p3...],
  //                      "Johnstown":[p1, p2, p3...], ...}

  var sortedAgeBySchool = {}
  for (var i = 0; i < peopleArr.length; i++) {
    // check for the school key
    if (!(peopleArr[i].school in sortedAgeBySchool)) {
      sortedAgeBySchool[peopleArr[i].school] = []
    }
    // now the school key is present
    sortedAgeBySchool[peopleArr[i].school].push(peopleArr[i])
  }

  // FIXME: Here is where I need a function to sort people with a 'Grouping'
  var peopleArrWithCalcRings = assignRingsMultGroups(sortedAgeBySchool)

  // now, populate
  // Populate the calculated ring in the source sheet
  for (var i = 0; i < peopleArrWithCalcRings.length; i++) {
    // col is calculatedRingCol + 1 to account for start at 0
    var cell = sourceSheet.getRange(
      peopleArrWithCalcRings[i].originalRow,
      peopleArrWithCalcRings[i].calculatedRingCol + 1
    )
    cell.setValue(peopleArrWithCalcRings[i].calcuatedRing)
  }
  console.log("Finished populating source page with calulated ring info.")
}

// Sort function for hashcodes.
function sortByNameHashcode(a, b) {
  return hashCode(a.sfn + a.sln) - hashCode(b.sfn + b.sln)
}

// sort function for height.
function sortByHeight(a, b) {
  return a.heightDec - b.heightDec
}

// Input: peopleArr, which is everyone in one grouping
//        startVRing, the place to start
//        numVRings, the number of rings to divide into
function divideOneGroupingIntoVRings(
  oneGroupingPeopleArr,
  startVRing,
  numVRings
) {
  groupDistArr = divideUpGroups(oneGroupingPeopleArr.length, numVRings)
  var vRingHash = {}
  var peopleArrIndex = 0
  for (i = 0; i < numVRings; i++) {
    vRing = i + startVRing
    var numPeopleInThisVRing = groupDistArr[i]
    for (j = 0; j < numPeopleInThisVRing; j++) {
      if (!(vRing in vRingHash)) {
        vRingHash[vRing] = []
      }
      vRingHash[vRing].push(oneGroupingPeopleArr[peopleArrIndex++])
    }
  }
  return vRingHash
}

// Given a peopleArr, divide into different Groups
// return: {"1": [p1, p2, ...], // the "1" is the group, so there could be discontinuities (don't rely on it being sequential)
//          "2": [p1, p2, ...],
//          "4": [p1, p2, ...] ...}
function separateIntoGroups(peopleArr) {
  // Find all the groups in the peopleArr

  var groupingsHash = {} // make it a hash
  for (var i = 0; i < peopleArr.length; i++) {
    thisGrouping = peopleArr[i].grouping

    if (!(thisGrouping in groupingsHash)) {
      groupingsHash[thisGrouping] = []
    }

    groupingsHash[thisGrouping].push(peopleArr[i])
  }
  return groupingsHash
}

// In: peopleArr
// out:   // sortedBySchool = {"Success": [p1, p2, p3...],
//                      "PA"     : [p1, p2, p3...],
//                      "Johnstown":[p1, p2, p3...], ...}
function divideBySchool(peopleArr) {
  var sortedBySchool = {}
  for (var i = 0; i < peopleArr.length; i++) {
    // check for the school key
    if (!(peopleArr[i].school in sortedBySchool)) {
      sortedBySchool[peopleArr[i].school] = []
    }
    // now the school key is present
    sortedBySchool[peopleArr[i].school].push(peopleArr[i])
  }

  return sortedBySchool
}

// Given the peopleArrOneGroup, order the people.
// Output is a peopleArr that is order based on the criteria.
function rankOneGroup(peopleArrOneGroup) {
  dividedBySchool = divideBySchool(peopleArrOneGroup)
  //dividedBySchool = {"Success": [p1, p2, p3...],
  //                   "PA"     : [p1, p2, p3...],
  //                   "Johnstown":[p1, p2, p3...], ...}
  // schools are unsorted right now

  // sort each school by age

  var dividedBySchoolSortedByAge = {}
  for (school in dividedBySchool) {
    dividedBySchoolSortedByAge[school] =
      dividedBySchool[school].sort(compareByAge)
  }

  // Now dividedBySchoolSortedByAge is the above, but the lists are sorted by age

  // Calculate the age rank
  for (school in dividedBySchoolSortedByAge) {
    numPeople = dividedBySchoolSortedByAge[school].length

    for (var i = 0; i < dividedBySchoolSortedByAge[school].length; i++) {
      // We want the first person on the list to have an age rank of 0. That way the first three will be guaranteed
      // to be from differnt schools as long as at least 3 schools are in the grouping.
      dividedBySchoolSortedByAge[school][i]["age rank"] = i / numPeople
    }
  }

  // Now, dividedBySchoolSortedByAge has 'age rank' field

  // Collapse it down to one peopleArr
  var rankedPeopleArr = []

  // Combine
  for (const value of Object.values(dividedBySchoolSortedByAge)) {
    rankedPeopleArr.push(...value)
  }
  // and sort by age rank
  rankedPeopleArr.sort(compareByAgeRank)

  return rankedPeopleArr
}

function assignRingsMultGroups(sortedAgeBySchool) {
  //sortedAgeBySchool is a hash of age-sorted people

  // Need to divide into kids, teen/adult men, and teen/adult women
  var dividedSortedAgeBySchool = { kids: [], men: [], women: [] } // will be 3 arrays, [0] is kids, [1] is men, [2] is women

  const kidMaxAge = 13
  for (var i = 0; i < sortedAgeBySchool; i++) {
    if (sortedAgeBySchool[i].age < kidMaxAge) {
      dividedSortedAgeBySchool["kids"].push(sortedAgeBySchool[i])
    } else {
      if (sortedAgeBySchool[i].gender == "M") {
        dividedSortedAgeBySchool["men"].push(sortedAgeBySchool[i])
      } else {
        dividedSortedAgeBySchool["women"].push(sortedAgeBySchool[i])
      }
    }
  }

  var startRing = 1
  for (var category in dividedSortedAgeBySchool) {
    var [peopleArrWithCalcRings, numRingsUsed] = assignRingsOneGroup(
      dividedSortedAgeBySchool[category],
      startRing
    )
    startRing += numRingsUsed
  }

  return peopleArrWithCalcRings
}

function assignRingsOneGroup(sortedAgeBySchool, startRing) {
  var rankedPeopleHash = applyAgeSchoolRank(sortedAgeBySchool)

  // Now, have to combine them all. Just mash them together and do a sort later.
  var combinedPeopleArr = []

  for (const value of Object.values(rankedPeopleHash)) {
    combinedPeopleArr.push(...value)
  }

  // Sort by the ageRank
  var sortedCombinedPeopleArr = combinedPeopleArr.sort(compareByAgeRank)
  console.log(
    "Finished combining all schools into one list and sorting by rank."
  )

  // choose the number of rings to use
  var numRings = Math.ceil(sortedCombinedPeopleArr.length / 10) // max 10 per ring

  var peopleArrWithCalcRings = putPeopleInRings(
    sortedCombinedPeopleArr,
    numRings,
    startRing
  )
  console.log("Finished assigning people to rings.")

  return [peopleArrWithCalcRings, numRings]
}

// Given a list of people in sorted order, place them into the number of rings specified
// Optionally give a starting ring number, default is 1
function putPeopleInRings(sortedPeopleArr, numRings, startRing = 1) {
  // Populate the calcuatedRing for each based on the number in each ring
  var groupArr = divideUpGroups(sortedPeopleArr.length, numRings)

  var overallPeopleArrIndex = 0

  for (var groupNum = 0; groupNum < groupArr.length; groupNum++) {
    // Get the number of participants in this ring
    var thisRingNum = groupNum + startRing
    var numPeopleInRing = groupArr[groupNum]
    // Now, pick off the next numPeopleInRing and assign them to this ring

    for (var j = 0; j < numPeopleInRing; j++) {
      sortedPeopleArr[overallPeopleArrIndex].calcuatedRing = thisRingNum
      overallPeopleArrIndex++
    }
  }

  return sortedPeopleArr
}

// Input: peopleArr
//        A sort function to use to sort them
//        A key to put in the output array
// Output: A peopleArr sorted by the sort function

function applySortOrder(peopleArr, sortFunction, sortKey, divideBySchoolFirst) {
  var unsortedPeopleWithRank = []
  var sortedPeopleWithRank = []

  // Divide by school if necessary
  if (divideBySchoolFirst) {
    // Divide by school
    var dividedBySchool = divideBySchool(peopleArr)

    // For each school, apply the sort function
    for (const [school, schoolArr] of Object.entries(dividedBySchool)) {
      var sortedSchoolArr = schoolArr.sort(sortFunction)

      for (var i = 0; i < sortedSchoolArr.length; i++) {
        sortedSchoolArr[i][sortKey] = i / sortedSchoolArr.length
      }

      unsortedPeopleWithRank = unsortedPeopleWithRank.concat(sortedSchoolArr)
    }
    sortedPeopleWithRank = unsortedPeopleWithRank.sort(function (a, b) {
      return a[sortKey] - b[sortKey]
    })
  } else {
    // Don't sort by school first, just sort the input arr based on the sort function
    sortedPeopleWithRank = peopleArr.sort(sortFunction)

    // Put the rank in the given sortkey for each person in the array
    for (var i in sortedPeopleWithRank) {
      sortedPeopleWithRank[sortKey] = i / sortedPeopleWithRank.length
    }
  }

  return sortedPeopleWithRank
}

// Input: peopleArr for people all in a virtual ring
// Output: another people array, with the 'Form Order' key filled in with the form order
function applyFormOrder(peopleArr) {
  // Get the list split into schools
  var dividedBySchool = divideBySchool(peopleArr)

  var unsortedPeopleWithRank = []
  // Sort each school by
  for (const [school, schoolArr] of Object.entries(dividedBySchool)) {
    var sortedSchoolArr = schoolArr.sort(function (a, b) {
      return hashCode(a.sfn + a.sln) - hashCode(b.sfn + b.sln)
    })
    for (var i = 0; i < sortedSchoolArr.length; i++) {
      sortedSchoolArr[i]["formRank"] = i / sortedSchoolArr.length
    }

    unsortedPeopleWithRank = unsortedPeopleWithRank.concat(sortedSchoolArr)
  }

  var sortedByFormRank = unsortedPeopleWithRank.sort(function (a, b) {
    return a.formRank - b.formRank
  })
  for (var i in sortedByFormRank) {
    sortedByFormRank["formOrder"] = i
  }

  return sortedByFormRank
}

/**
 * Returns a hash code for a string.
 * (Compatible to Java's String.hashCode())
 *
 * The hash code for a string object is computed as
 *     s[0]*31^(n-1) + s[1]*31^(n-2) + ... + s[n-1]
 * using number arithmetic, where s[i] is the i th character
 * of the given string, n is the length of the string,
 * and ^ indicates exponentiation.
 * (The hash value of the empty string is zero.)
 *
 * @param {string} s a string
 * @return {number} a hash code value for the given string.
 */
hashCode = function (s) {
  var h = 0,
    l = s.length,
    i = 0
  if (l > 0) while (i < l) h = ((h << 5) - h + s.charCodeAt(i++)) | 0
  return h
}
// Apply the Age/School ranking system.
function applyAgeSchoolRank(sortedAgeBySchool) {
  var schools = Object.keys(sortedAgeBySchool)

  console.log(schools)
  var tmpSchoolAgeSorted = {}

  for (var i = 0; i < schools.length; i++) {
    // sort one school
    schoolSortedByAge = sortedAgeBySchool[schools[i]].sort(compareByAge)
    // Add ranking here
    totalPeople = schoolSortedByAge.length
    for (var j = 0; j < schoolSortedByAge.length; j++) {
      schoolSortedByAge[j]["age rank"] = (j + 1) / totalPeople
    }
    // Add to a new table
    tmpSchoolAgeSorted[schools[i]] = schoolSortedByAge
  }

  //  console.log(tmpSchoolAgeSorted);

  return tmpSchoolAgeSorted
}

// Entry point. Do all the levels, or pick one to do.
function generateOverview(
  level = false,
  readFromCalcRings = false,
  useRemapping = false
) {
  var thisSpreadSheet = currentSpreadsheet()
  const catArr = globalVariables().levels

  if (level) {
    printRingsOneLevel(level, readFromCalcRings, (useRemapping = useRemapping))
  } else {
    for (var i = 0; i < catArr.length; i++) {
      printRingsOneLevel(
        catArr[i],
        readFromCalcRings,
        (useRemapping = useRemapping)
      )
    }
  }
}

// Translates a ring number into a physical ring location
// return: string ring, x, and y (x and y are 1-based)
function physicalRing(ring, totalRings, numPhysicalRings) {
  var retPhysRing
  var doubleUps = totalRings - numPhysicalRings

  var x
  var y

  if (totalRings <= numPhysicalRings) {
    retPhysRing = ring
    x = ring
    y = 1
  } else {
    if (ring <= doubleUps * 2) {
      var phyRing = Math.floor((ring + 1) / 2)
      if (ring % 2 == 1) {
        retPhysRing = phyRing + "a"
        x = phyRing
        y = 1
      } else {
        retPhysRing = phyRing + "b"
        x = phyRing
        y = 2
      }
    } else {
      retPhysRing = ring - doubleUps
      x = retPhysRing
      y = 1
    }
  }

  return [retPhysRing, x, y]
}
