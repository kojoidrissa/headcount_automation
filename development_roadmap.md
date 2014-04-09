#Development Roadmap

##headcountSummary.py
-  make `worksheet_to_table` function more robust with Glenn Z.'s suggestions:
    -  Use column/header names to ID the source data columns
        -  I'll probably need to create a tuple of strings to contain the headers; then the code can cycle through those & find their indices
    -  Include code that will throw a VISIBLE exception if a needed column is missing
        -  v. 1: print to StdOut; stop the program; log the exception somewhere (a log file?)
        -  v. 2: create an Alert box in the GUI; log the exception somewhere (a log file?)

##functionalSheets.py
-  ALL output from `funcSheets_check_figures` function  needs to be captured in a file, not just printed on screen
    -  Do I make this a log file or just add it to the Exceptions spreadsheet?
        -  The log file would be a useful exercise for me. It could be a text file that lives in that directory with a date of when it was run.

##format.py
-  Since my headers AND footers are now bolding (as of me running the  code in April of 2014; no idea **why**), I can start looking at applying number formats to the hours and % columns.

