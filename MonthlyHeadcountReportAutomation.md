#Monthly Headcount Automation Project

##Current, mostly manual process
Currently(7/10/2013 10:52:15 AM ), we follow these steps to create the Headcount Reports each month:

1. Get Kronos headcount reports from Hilton (for 1008, 1902 & 2231)
2. Combine data from these reports into a single document as the basis for the *Raw Data* document
3. Supplement *Raw Data* with information from *Weekly Cost Center* reports (from Payroll) and *Cost Centers & Managers* spreadsheet (which I maintain)
    -  Details on this process are in *WorkJournal.md*

The MAIN purpose of this *Raw Data* sheet is to allow the creation of a pivot table to summarize it's data in a way that's:

1. Easy to read (although no one reads it)
2. Easy to copy/paste into the Monthly Hdcnt Summary sheet and the various Functional Area sheets, which are also formatted for easy reading. We use THESE in our monthly meetings with various managers (Mason, Whitcomb, etc)

###Primary Inputs
-  Kronos Headcount Reports (for each Entity)
-  Weekly Cost Center Reports (perhaps 2, to bookend the month)
-  Cost Centers & Managers spreadsheet; I need to:
    -  Find a canonical home for this on a shared drive
    -  Come up with a better update strategy

###Primary Outputs
-  Raw Data Sheet <del>& Pivot Table</del>
-  Monthly Hdct Summary sheet
-  Functional Area headcount reports (Engineering, Subsea, Project Controls, etc.)


##Plans for automated process
The ultimate goal is to be able to take the primary inputs listed above and have the required outputs pop out the other side. That will likely have to happen in stages.

###Version 0.25: Raw Data to Hdct Summary: Pivot-free MVP (*DONE*)
Once the *Raw Data* sheet has been created, it shouldn't be too difficult to generate the Functional Tabs. But I'll start with the Monthly Hdct Summary first. It's a simpler, more general case of the Functional tabs. This also lets me avoid the problem of creating the list of Cost Centers that need to make up each Functional tab.
####Simplyfying decisions for this phase
-  Hardcoding the column positions to use for Key construction. I'll get 'general' in another iteration
-  "Manual" insertion of `CC Description`; I'll figure out how to pull it in from *Cost Centers & Managers* later. Maybe that'll be Version 0.75
-  <del>Ignore the total DOE & Project hours for now. That'll be for Version 0.5</del>
####Component Parts
-  Can I calculate a unique key and store it as part of a data structure I can write?
    -  Yes
 
            d ={}
            x ='1008'
            y = '57230'
            z = '160470'
            d[(x+y+z)] = 7

            d
            #Out[184]: {'100857230160470': 7}
-  Indexes for info needed 
    -  **while** reading from Raw Data
        -  Employee Num: 0
        -  CostCenter: 2
        -  Company: 3
        -  Employee Name: 10
        -  Manager: 4
        -  DOE or Project: 17
        -  Total Hours: 21
    -  **after** already read in
        -  Employee Num 0
        -  CC 1
        -  Company 2
        -  Manager 3
        -  Employee Name 4
        -  DOE/Project 5
        -  Total Hours 6

-  <a id ='dict'>Basic data structure</a> to create for writing

    If a[0] is the header list, the output order needed is

        dict = {CompanyCostCenterEmpnum: ['Company', 'CC', 'EmpNum', 'EmpName', 'Manager', {'DOE' : DOEHrs, 'Project' : ProjHrs}]}
    So, for the version without DOE/Proj Hours, we'd have

        dictHeader = {a[0][2]+a[0][1]+a[0][0] : (a[0][2], a[0][1], a[0][0], a[0][4], a[0][3])}
    -  Need to remember how to access JUST the keys from a dictionary
        -  d.keys() returns the keys in a dictionary
        -  d.iterkeys() returns "*an iterator over the mapping's keys*"
            -  I'm not **100%** sure what that means, but I'm PRETTY sure "*an iterator*" is what I need when I was trying to iterate over a sequence (or a mapping, in this case) and I'm told "An index needs to be an integer, not a list". This is what caused me to use the "range(len(table))" construct, to get integers for the number of items in *table*
            -  In retrospect, this was probably caused by my using nested tables, as each item INSIDE the list, *table*, was another list
        -  <del>`dictionary.items()[0][0]` returns the first key from the first pair in a dictionary</del> **THIS IS A HACK! NOT THE PROPER WAY TO GET *DICTIONARY* KEYS!!**



###Version 0.5 (*DONE*)
Compute total DOE & Project hours for a constructed key. That'll be for Version 0.5

###Version 0.75
-  Converting all the data in the columns that AREN'T Hours to the type (Unicode). Right now the Company, Employee Num and some of the CCs are coming through as Int.

-  Create Functional Sheets Module

    This phase will probably have the list of cost centers for each Functional area hard-coded, as a set of lists or dictionaries. The longer-term goal is to have this list live in a plain text config file of some sort. That way, changes (I can imagine the 2231/Subsea cost centers switching tabs, as well as other adjustments) can be made easily, without needing to edit (or be able to read/understand) code.

    -  Do I want this to be part of the same file/module(headcount.py), or to be a separate module that uses the output of 'headcount.py'?
        -  A separate module would reduce the risks of accidentally 'breaking' *headcount.py*. It would also, I think, help make the code more modular. These are two different problems.
        -  So, this isn't really Ver 0.75 of this code, but of the entire PROJECT. A second module of code to add the 0.75 functionality
        -  This module should take the sheet created by headcount.py and:
            -  be able to deal with it, no matter how many columns it has (with or w/o the CC Descriptions added)
            -  Create a section for each Company/CC combo, with a space between each group. 
                -  Can I use 'Named Ranges' here?
                -  Will I need to create a different "key" for each unique combo? I'm thinking about the subsea cost centers, that sometimes are used in 1902 and 2231. The code should be able to handle any other circumstances like that which I HAVEN'T forseen.
            -  Create a space b/w each section

    #### Simplyfying assumptions
    -  Leave out the automatically calculated section totals for now
    -  No formatting (borders, etc)

###Version 1.0: Raw Data to Functional Tabs: Pivot-free
"Manual" insertion of `CC Description`; I'll figure out how to pull it in from *Cost Centers & Managers* later. Maybe that'll be Version 0.75

####Completion steps for this version
Needed for each sheet

-  Insert "Total Hours", "DOE Util. %" and "Proj. Util. %" column headings; also insert the formula into the first row
-  Fill in the "Total Hours", "DOE Util. %" and "Proj. Util. %" formulae
-  Mark whole sheet for filtering
-  Insert two columns b/w 'CC' and 'Employee Num'
    -  CC
    -  CC Description
-  Convert the CC numbers into text (`=TEXT(B2,"#")`) and paste values those amounts in the first of the two added columns; include the 'CC' rows in your filter
    -  This is a stopgap measure. I need to figure out how to have the CC values come out AS strings, not numbers
-  With the CC numbers in place, `VLookup` the Cost Center Descriptions
-  Section totals
    -  Do the managers WANT/NEED these?
    -  Yes; add to code
-  Accounting format the numbers

####Speed bottlenecks: first problems to address
-  Insertion of "Total Hours", "DOE Util. %" and "Proj. Util. %" column headings & formula
    -  The column headings aren't an issue. I just need to figure out how to do the formulae       
-  Insertion of CC Descriptions
    -  This is the slowest part of the manual process, but still useful. Most managers don't know their Cost Center numbers. Neither do I!

####Style Guidelines for ALL sheets
-  Utilization columns: % formatted, no decimals
-  Hours columns: Accounting formatted, no decimals
-  Cost Center Headers & footers: **bold**
-  Align TEXT cells Left-Aligned; HOURS cells Right-Aligned
-  Print Optimization: Last Step
    -  Page Header: *Month* **Year** Headcount Utilization Report; Tab
    -  Page Footer: Same as above, with "*#/##*" in right column
    -  Page Orientation: Landscape

###Version 1.0.1 (for running August reports in early September 2013)

####Completion steps for this version
1. Get Kronos headcount reports from Hilton (for 1008, 1902 & 2231)
2. Combine data from these reports into a single document as the basis for the *Raw Data* document
3. Supplement *Raw Data* with information from *Weekly Cost Center* reports (from Payroll) and *Cost Centers & Managers* spreadsheet (which I maintain)
-  Create the 5 extra columns in Raw Data
-  After Hdcnt Summary is created: 
    -  Copy that tab into a new sheet, rename the original to "MHS Original"; name the copy "Monthly Headcount Summary" and work on that 
    -  Insert two columns b/w 'CC' and 'Employee Num'
        -  CC
        -  CC Description
    -  Convert the CC numbers into text (`=TEXT(B2,"#")`) and 'paste values' those results in the first of the two added columns; include the 'CC' rows in your filter
        -  This is a stopgap measure. I need to figure out how to have the CC values come out AS strings, not numbers
    -  With the CC numbers in place, `VLookup` the Cost Center Descriptions and 'paste values' those results
-  Use this updated version of Hdcnt Summary as the input for functionalSheets.py
    -  It looks like I'll need to change `headcount_sorted = sorted(functable, key = itemgetter(1, 0, 3))` in the *create_tabs* function. The last index will need to be 4, since I'm adding the CC Description column.
        -  I ALSO needed to adjust the size of the Footer rows to be one longer. That led me to also make the spacer rows longer AND to figure out how to generate each of those lists programmatically, instead of manually. List comprehensions (and [StackOverflow](http://stackoverflow.com/a/10712032)) to the resucue! 
        -  This also highlights the need for me to use header names to generate indcies, instead of hard-coding them. It makes the process too fragile.
    -  If it doesn't work immediately, I'll have to go back to the manual method of adding in CC Descriptions AFTER functionalSheets.py does it's work


-  Version 1.0.1 results: Headcount

        24.8499999046
        Loading time for <Worksheet "raw data"> :  24.8499999046
        Time to create 'Table' from  <Worksheet "raw data"> for 11718 rows and  22 columns:  1.01399993896
        Loading time for <Worksheet "raw data">  : 24.8499999046
        11718 Rows;  22 Columns
        durTable 1.01399993896
        11718 Rows;  7 Columns
        durKeylist 0.296999931335
        durHourlist 1.88700008392
        durFinalTable 2.74599981308
        durFinalTableMem 0.0620000362396
        Writing time for hdcntsum.xlsx  : 0.344000101089
        445 Rows;  7 Columns
        durTotal 32.1510000229  

-  Version 1.0.1 results: functionalSheets

        fullTable Creation Time was  43.9909999371 seconds.
        Creation Time  for all Functional Tables was  67.625 seconds.
        Creation Time for ALL tabs was  0.25 seconds.
        Total processing time 113.674999952



####Changes/Updates meant for this version (this may happen AFTER the reports are complete)
-  **Create a Footer for each Cost Center, with CC totals for DOE, Project, Total Hours, and Utilization %s** 
    -  Creating a separate function to generate Footer rows for each Cost Center (from a completed funcTable) seems to make the most sense
        -  I can insert that Footer row in place of the first "spacer" row in the *create_tabs* function: `ws.append(spacer) #One for summation of the section above`
    -  from within each Cost Center grouping, `makeSubseaTable` and `makeNoSubseaTable`
        -  after temptable is complete, add CC total footer row? Or create a separate function to process a completed temptable?
-  **Calculate Utilization %s for the entire Functional Area**
    -  I'll do this after the August report is done:
        -  costCenterFooter.py probably needs to be a **module** with multiple functions:
            -  One that returns ccUnique
            -  One that returns footer
            I should be able to use those two results to easily calculate a footer for the entire functional area. But I don't really have time to work on that now (2013-09-05 09:13); I'll work on that after I finish the August reports. Probably on 2013-09-06.


-  Optimized reading/writing: 
    -  Compare value of "Optimized" vs. reading/writing by row/column
-  Proper use of dictionaries instead of nested lists
    -  I'd planned this in my original design, but fell back on Lists due to time constraints, since I was more familiar with them. See [Basic data structure to create for writing](#dict)

###Version 2: Automating Raw Data creation
-  Do I need to add the Five Left Columns (*used to get Company, CC and truncated Employee Number*) to Hilton's file before compressing it into the Hdcnt Summary? That workflow MIGHT be artifact of the Excel-based process.
-  I can probably create the Hdcnt Summary FIRST (*turning ~12K rows of data into ~500*), **then** add in the other columns via EmpNum lookup. Doing it on ~500 rows should be WAY faster than on ~12K rows.
-  **Figure out how to use Hilton's Kronos file in more of a "raw" format**
    -  This may require a larger-scale refactoring, as it would change how things are processed from the very beginning.
    -  having code pull in the manager names and cost centers from the payroll files
    -  letting code split the Entity/Cost Center info
    -  Cutting the leading "00" from the Employee Number

###Version 3: Functional Tab Config File

###Version 4: One Program To Rule Them All! Input Files ==> Reports

##Refactor notes

###Selecting columns by Header Name, not Index
In the following code block, from headcountJuly2013.py:

    for row in range(len(source.rows)):
        r =[]
        #adding 'ref' selection tuple dropped loop time on 89 cols from 0.416000 sec to 0.008000
        #reordering the tuple indexs here to avoid having to shuffle them later
        #This will be less fragile if I take the following advice from Glen:
            #do this by column header/name instead of index
            #include code that will throw a VISIBLE exception if a needed column is missing
        ref = (3, 2, 0, 10, 4, 17, 21) 
        for col in ref: #Original "in" argument was 'range(len(source.columns))'
            r.append(source.cell(row = row, column = col).value)
        table.append(r)
    end = time.time() #End Timer
    durTable = end - start #with 1,000 rows, durLoop = datetime.timedelta(0, 47, 16000) or 47 seconds
    
The `ref = (3, 2, 0, 10, 4, 17, 21)` line needs to be replaced with code that does the following:

-  Finds the required strings in the header row
-  returns those header indexes in the form of some sort of iterable (not sure what, exactly; would a tuple be better than a list?)
-  I should probably keeps that sequence of header strings, since I'll want to refer to them later

###Running on the Network Drive
From within the Canopy shell:

`cd 'M:\Dbsteam\BUDGET\Jackie\MNTH_RPT\2013\<MMMM> <YYYY>\Headcount Misc'`
