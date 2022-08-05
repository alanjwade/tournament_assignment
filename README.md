# tournament_assignment
Using a template in Google Sheets, take a list of people with some attributes (age, height, whether they are doing forms or sparring), and generate a forms ring assingment and a sparring ring bracket.
## General goals
1. Have a 'calculate' step. This takes the data and puts people into rings based on human input.
2. 'tweak' step. This takes the calculated values and lets a person move people around to their liking.
3. 'print' step. This creates the checking sheets, as well as the forms and sparring scoring sheets.

Steps 2 and 3 can be run multiple times. Step 1, after the initial time, should not be run again since it will wipe away any later tweaks done in step 2.

## Process
1. Make the registration spreadsheet with the following columns (exact match, case counts):
    * Student Last Name
    * Student First Name
    * School
    * Age
    * "Height (feet)"
    * "Height (inches)"
2. Add these columns.
    * "Grouping"
      * In here, put the same number for every person that are at the same ability level. For example, maybe you want 9 and 10 year olds together, so you'd give them all a grouping of 4 (assuming you had groupings 1-3 already assigned for younger kids). Technically the actual number itself isn't significant, but it may help to keep the rings in ascending age order.
    * "Virtual Ring"
      * Leave this blank. The script will fill in this column, although you will be able to tweak it later.
    * "Form Order"
    * "Sparring Order"
3. Run the 'Assign Virtual Rings' step from the menu.
    * This will fill in the "Virtual Rings", "Form Order", and "Sparring Order" columns. It will use an algorithm to try to not have 3 people from the same school first for forms, and it will try to keep the first round of sparring heights 'close' and not have people from the same school spar as much as possible (in the first round).
4. Fill in the 'Virtual to Physical Mapping' table on the registration sheet.
5. Adjust the 'Virtual Ring', 'Form Order", and "Sparring Order" columns to make changes.
6. repeat steps 3 through 5 until satisfied. Run 'Generate overview' to get the sheet that has all the ring assignments on it.
7. Run the 'Generate forms scoresheet', 'Generate sparring scoresheet', 'Generate checkin sheet' to get the final deliverables.
8. Print hard copies of the sheets.


## How ring assignment is determined
  * Students are divided up be school. Within the school, they are ordered by age.
  * They are then given a fractional number based on their order with their school
    * if 3 people in the school, the numbers would be 0, .333, and .666
    * if 4 people, 0, .25, .5, .75
  * Then, all the students are put back into a single list
  * This list is then divided up into how many rings are needed for the particular grouping.
## How form order is determined
  * Inside a particular ring, the names are hashed and sorted.
## How sparring order is determined
  * Inside a particular ring, the people are ordered by height.