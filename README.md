# Example project to illustrate an issue
## How To Use
- Load the solution in VS
- Run the project
- Look in the output folder (`./bin/debug`)
- Open the .xlsx file that was produced
- Note that Cell B2 is locked, even though it is explicitly set as unlocked in the code

Notes:
The code is contrived example based on how our production code constructs a spreadsheet.
Removing this line:
``
    sheet.Rows[0].SetIsLocked(false);
``
 apprears to fix the issue, but that is not possible in our production code :-/
 