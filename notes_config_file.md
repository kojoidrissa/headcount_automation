#`config_file`  branch planning notes

If I'm going to use `costCenter_Function_map.json` as my config file, that'll require some changes to `functionalSheets.py`. I should be able to replace the two big blocks running functions with calls that reference the .json dictionary. But HOW exactly will I do this? **NOTE:**  The current config file is accurate as of 2013-08-01; needs to be updated for 2013-09 reports

##Block for Function 1: functionTable ==> makeSubSeaTable/makeNoSubseaTable
From the function definition:

<pre>
    ''' 
    list ==> list of lists
    
    Input a list of cost centers for a functional tab;
    return a list of lists where each inner list is a row of representing one person in that Cost Center
    '''
</pre>

I should be able to replace this block:
<pre>
    subsTable = makeSubseaTable(subsea)
    estSTable = makeNoSubseaTable(estSales)
    prjCTable = makeNoSubseaTable(prjCntrls)
    infoTable = makeNoSubseaTable(infotech)
    procTable = makeNoSubseaTable(procure)
    legaTable = makeNoSubseaTable(legal)
    engiTable = makeNoSubseaTable(engineering)
    humaTable = makeNoSubseaTable(humanres)
    prjMTable = makeNoSubseaTable(prjMgmt)
    accoTable = makeNoSubseaTable(accounting)
    ethiTable = makeNoSubseaTable(ethics)
    hsesTable = makeNoSubseaTable(hses)
    qualTable = makeNoSubseaTable(quality)
</pre>
with
<pre>
    import json
    dept_dict = json.load(file('costCenter_Function_map.json'))
    sheet_dict = {}
    for key in dept_dict.keys():
        if key != 'Subsea':
            sheet_dict.update({key : makeNoSubseaTable(dept_dict[key])})
        else:
            sheet_dict.update({key : makeSubseaTable(dept_dict[key])})
</pre>

## Block for Function 2: create_tabs
from the function definition:
<pre>
    '''
    list of lists, string --> list of lists

    Takes in nested list for each functional area (created by funcTable/makeSubseaTable/makeNoSubseaTable) and a string (tabname)
    each inner list reperesents a row of data; creates a spreadsheet in memory, writes those rows to the spreadsheet
    The string becomes the name of the worksheet
    '''
    </pre>

The code block is:
<pre>
    create_tabs(subsTable, 'Subsea')
    create_tabs(estSTable, 'Estimation & Sales')
    create_tabs(prjCTable, 'Project Controls')
    create_tabs(infoTable, 'IT Services')
    create_tabs(procTable, 'Procurement')
    create_tabs(legaTable, 'Legal')
    create_tabs(engiTable, 'Engineering')
    create_tabs(humaTable, 'Human Resources')
    create_tabs(prjMTable, 'Project Management')
    create_tabs(accoTable, 'Accounting')
    create_tabs(ethiTable, 'Ethics')
    create_tabs(hsesTable, 'HSES')
    create_tabs(qualTable, 'Quality')
    create_tabs(fullTable, 'Headcount Summary Sorted')
</pre>
If I have Block 1 create a 2nd dictionary with the format {'Dept1Name' : [NestedDept1Table], 'Dept2Name' : [NestedDept2Table]... }, I **SHOULD** be able to replace this block with:
for key in sheet_dict.keys():
    create_tabs(sheet_dict[key], key)
create_tabs(fullTable, 'Headcount Summary Sorted')


