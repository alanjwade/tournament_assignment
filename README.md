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
3. Run the 'Calculate' step from the menu.
4. Adjust the 'Virtual Ring' column to make changes.
5. repeat steps 3 and 4.