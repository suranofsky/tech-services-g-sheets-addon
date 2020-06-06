This project (MatchMARC) is a Google Sheets Add-on created for the Lehigh University Libraries technical services team.
<br><br>
I collaborated on this project with Lehigh's Cataloging/Metadata Librarian, Lisa McColl, who helped me understand the needed functionality.  She did all of the testing and provided feedback for each iteration of the project.
<br><br>
We co-presented a session about this project to the ALA Technical Services Workflow Efficiency Interest group during the American Library Association's 2019 annual conference. You can view our slides here: https://connect.ala.org/HigherLogic/System/DownloadDocumentFile.ashx?DocumentFileKey=3a69473d-a4a3-4781-a546-72b394ef3886
<br><br>
We are co-authoring an article about the project which will be published in the November 2019 issue of the code4lib journal.  
<br><br>
This add-on is publicly available using the “Add-ons > Get add-ons” menu in any Google Sheet. A search for MARC will show the add-on.
<br><br>
You can also find the add-on here:
https://gsuite.google.com/marketplace/app/matchmarc/903511321480


# Versions

### 6-6-2020
### Version 12
## New Features and Fixes
1) If it finds duplicate local holdings, it will select one and bold the row so you know it found a duplicate
2) If no match is found, it will select the top match (with the most holdings) only if you check the box: Select first record when no match?
3) New ISSN Search: If you label row one, column one "ISBN" it will search by ISBN.  If you label row one, column one "ISSN" it will search by ISSN

### 2-7-2020
### Version 11
## New Features:
1) Email is now separate from the search/match functionality.  After you have searched for the records, the 'email' is a 2nd step.  The email will use the 001 values found in the search to retreive the MARC records.  You will have to fill in the field informing the script which column the 001 value is in.
2) On the search criteria tab, you can only configure the *starting* column for the data you want written to the spreadsheet. This was done to speed up the execution.  Google Apps script is much faster when it is writing a block of cells instead of writing one cell at a time.  This speeds up the execution considerably.
3) When you email the file to yourself, you can add MARC fields to each record.  Configure this by putting the field/subfield in the column heading and the values in the rows for each record.  You can add multiple fields to each record.  (Please see screen print below for clarification)
4) You can have it write up to 25 fields to the spreadsheet.  
5) If you request an emailed file, the file will exclude duplicate records


![Illustrates the new feature - add fields to MARC record that will be emailed](matchMarchScreenShot.png?raw=true "Illustrates the new feature - add fields to MARC record that will be emailed")

