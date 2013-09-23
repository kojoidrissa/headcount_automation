#!python

##cd 'C:\Users\kidrissa\Documents\Monthly Headcount Schedule\July 2013 Headcount'
##This will take the output of "headcount.py" and create the various functional spreadsheet summaries
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import time

start = time.time() #timer, until I learn to use http://docs.python.org/2/library/profile.html
wb = load_workbook('hdcntsum.xlsx')
source = wb.get_sheet_by_name('Monthly Headcount Summary')

#creating the header row
header =[] 
for c in source.rows[0]:
    header.append(c.value)
#I STILL don't know why those 0s got added to the header; need to figure that out
#Probably because the source data had ONE 'DOE/Project' column, not two separate cols
header[-2:] = 'DOE', 'Project'
header.extend(['Tot. Hours', 'DOE Util %', 'Proj. Util %']) 

##Lists of Cost Center codes 'CC'that make up each functional tab
#Should these be dictionaries with {function: [cc1, cc2,...ccn]}? Probably.

engineering = ['57100', '57240', '57250', '57260', '57620','57166']
#subsea will use ALL 57230, but ONLY use the other Cost Centers where they're in Company 2231
#57619, 61201, 61316, 61708 are all used outside of 2231
subsea = ['57230','57621', '57619', '61201', '61316', '61708']
prjCntrls = ['55221', '57171', '52A04']
#prjMgmt also has Cost Centers that cross Company lines
prjMgmt = ['52A02', '52P01', '52P02', '52P04', '52P05', '52P06', '52P07', '52P08', '52P09', '52P10', '52P11', '52P14','52P17', '52P18', '52P19']
estSales = ['55831', '52215', '61820', '61708', '61745']
quality = ['52A03', '51346', '55054', '57165']
accounting = ['61101', '57619', '61160', '61174']
humanres = ['55832', '55833', '61316', '55834']
infotech = ['61201']
hses = ['51247', '51270', '55308', '55358', '52A01']
legal = ['61401']
procure = ['55236', '52A05']
ethics = ['61315']

#I'm not 100% sure how to optimally use this dictionary, so I'm going to go with a less optimzed approach
#I'm going to just repeat the code. (2013-07-22)
#However, I AM going to use this construct to generate a list of the different tabs I need.
#REVISIT THIS DICTIONARY (2013-07-23)
#ALSO: I made this dictionary by coverting the Cost Center lists above into a dictionary. Why? Because I'd created the above lists first
#I need to revisit this dictionary. There should be some way for me to feed THIS
'''
tab_dict = dict(engineering = ['57100', '57240', '57250', '57260', '57620', '57166'], subsea = ['57230', '57619', '57621', '61201', '61316', '61708'],
prjCntrls = ['55221', '57171', '52A04'],
prjMgmt = ['52A02', '52P01', '52P02', '52P04', '52P05', '52P06', '52P07', '52P08', '52P09', '52P10', '52P11', '52P14','52P17', '52P18', '52P19'],
 estSales = ['55831', '52215', '61820', '61708', '61745'], quality = ['52A03', '51346', '55054', '57165'], accounting = ['61101', '57619', '61160', '61174'],
 humanres = ['55832', '55833', '61316', '55834'], infotech = ['61201'], hses = ['51247', '51270', '55308', '55358', '52A01']
, legal = ['61401'], procure = ['55236'], ethics = ['61315'])
tablist = tab_dict.keys()
'''

#This function takes in the list of cost centers which make up a Functional Tab, returns a list of lists; 
#each internal list is a row for that Functional tab
##I'd originally intended it to return a list of tuples; why the change? Probably lack of skill on my part
##I could tuple() each temprow BEFORE I append it to tempTable; or learn how to build tuples programmatically
##The second seems the best option; one or both will have to happen in a later refactor
###As of 2013-08-01, This function is being cloned into two different functions;One that excludes Company 2231(Subsea)
###and another that excludes anything NOT Company 2231(Subsea) This is my current (2013-08-01)
###solution to the issue of Subsea having cost centers used in 2231 and 1902 in a Cost Center-focused report
###Also, I need to get Git installed on this machine

#def functionTable(list):
#    
#    ''' 
#    list ==> list of lists
#    
#    Input a list of cost centers for a functional tab;
#    return a list of lists where each inner list is a row of representing one peron in that Cost Center
#    '''
#    
#    tempTable = []
#    for row in source.rows:
#        ri = source.rows.index(row)
#        if str(row[1].value) in list:
#            temprow = []
#            for cell in row:
#                ci = source.rows[ri].index(cell)
#                temprow.append(source.rows[ri][ci].value)
#            tempTable.append(temprow)
#    return tempTable

##This function is intended to create a table of the ENTIRE contents of the headcount summary file
fullTable = []
time1 = time.time()
for row in source.rows[1:-1]: #trying to work around problem with 1st and last rows; was 'for row in source.rows'
    ri = source.rows.index(row)
    temprow = []
    for cell in row:
        ci = source.rows[ri].index(cell)
        #print source.rows[ri][ci].value
        temprow.append(source.rows[ri][ci].value)
        #print temprow #only for debugging purposes; I want to see when/where the float division by zero problem is happening
        #It seems to be in the "hdcntsum.xlsx" header row, where the DOE & Project columns have '0' in them.
        #I'm taking a slice of source.rows to ignore that first row for now
    temprow.append(temprow[-1]+temprow[-2]) #Total Hours: sum of DOE & Proj Hours
    temprow.append(temprow[-3]/float(temprow[-1])) #DOE Util%; DOE Hours / newly added Total
    temprow.append(temprow[-3]/float(temprow[-2])) #Proj Util%; Proj. Hours / Total
    fullTable.append(temprow)
time2 = time.time()
print "fullTable Creation Time was ", time2-time1, "seconds."

#This function takes in the list of cost centers in Subsea, returns a list of lists; 
#each internal list is a row for Subsea. It also excludes any rows that aren't
#Subsea (Company 2231)
def makeSubseaTable(list):
    
    ''' 
    list ==> list of lists
    
    Input a list of cost centers for a functional tab;
    return a list of lists where each inner list is a row of representing one peron in that Cost Center
    '''

    tempTable = []
    for row in source.rows:
        ri = source.rows.index(row)
        if row[0].value == 2231: #this is what was missing. I need to test the VALUE of what's in that cell
            if str(row[1].value) in list: #row[1] is the position of the Cost Center; I should change that to get the index of the name
                temprow = []
                for cell in row:
                    ci = source.rows[ri].index(cell)
                    temprow.append(source.rows[ri][ci].value)
                temprow.append(temprow[-1]+temprow[-2]) #Total Hours: sum of DOE & Proj Hours
                temprow.append(temprow[-3]/float(temprow[-1])) #DOE Util%; DOE Hours / newly added Total
                temprow.append(temprow[-3]/float(temprow[-2])) #Proj Util%; Proj. Hours / Total
                tempTable.append(temprow)
    return tempTable

#This function takes in the list of cost centers in Subsea, returns a list of lists; 
#each internal list is a row for Subsea. It also excludes any rows that aren't
#Subsea (Company 2231)
def makeNoSubseaTable(list):
    
    ''' 
    list ==> list of lists
    
    Input a list of cost centers for a functional tab;
    return a list of lists where each inner list is a row of representing one peron in that Cost Center
    '''

    tempTable = []
    for row in source.rows:
        ri = source.rows.index(row)
        if row[0].value != 2231: #this is what was missing. I need to test the VALUE of what's in that cell
            if str(row[1].value) in list: #row[1] is the position of the Cost Center; I should change that to get the index of the name
                temprow = []
                for cell in row:
                    ci = source.rows[ri].index(cell)
                    temprow.append(source.rows[ri][ci].value)
                temprow.append(temprow[-1]+temprow[-2]) #Total Hours: sum of DOE & Proj Hours
                temprow.append(temprow[-3]/float(temprow[-1])) #DOE Util%; DOE Hours / newly added Total
                temprow.append(temprow[-3]/float(temprow[-2])) #Proj Util%; Proj. Hours / Total
                tempTable.append(temprow)
    return tempTable
    
    
##Creating the target workbook
target = Workbook()
dest_filename = r'hdcntfunc.xlsx'

##Function that creates the individual functional area worksheets IN the workbook just created
def create_tabs(functable, tabname):
    '''
    list of lists, string --> list of lists
    
    Takes in nested list for each functional area (created by funcTable/makeSubseaTable/makeNoSubseaTable) and a string (tabname)
    each inner list reperesents a row of data; creates a spreadsheet in memory, writes those rows to the spreadsheet
    The string becomes the name of the worksheet
    '''
    #creating & naming the spreadsheet in memory
    ws = target.create_sheet(0)
    ws.title = tabname
    #ws.append(header) THIS was a mistake.
    
    #sorted_by_second sorts the incoming data by the 2nd element in each list; the Cost Center in this case
    #thanks StackOverflow! http://stackoverflow.com/questions/3121979/how-to-sort-list-tuple-of-lists-tuples
    #sorted_by_second = sorted(functable, key=lambda tup: tup[1])
    
    ##I upgraded to using the Operator Module, as it lets me sort by MULTIPLE criteria; 
    ##http://docs.python.org/2/howto/sorting.html#operator-module-functions
    from operator import itemgetter
    ##using Operator module, sorting by Cost Center, then Company, then Emp. Name
    ###I later created a function, sort_criteria, to generate the itemgetter keys based on the header values

    import sortCriteria.sort_criteria
    sort_by = sort_criteria(source.rows[0])
    headcount_sorted = sorted(functable, key = itemgetter(1, 0, 4)) #changed 3rd index from '3' to '4' due to change in Hdcnt Summary
    
    #goes through the sorted nested list, writing it to the spreadsheet in memory
    #Updated version uses the APPEND method from http://pythonhosted.org/openpyxl/api.html#module-openpyxl-worksheet-worksheet
    
    #spacer added to create break for manual insertion of Cost Center sum functions
    spacer = [None for i in range(len(functable[0]))] #used range in a list comprehension to build this

    #bring in my custom Footer code and generate the Footer dictionary for this Functional area
    ##since I didn't adjust the Path variable (I'll do that later), costCenterFooter.py had to
    ##be in the same directory as the files it was working on
    ###Forgot I need to do moduleName.functionName() when calling a func from a module
    ###Since I named my module AND function the same thing (I hadn't planned on it being a module)
    ###I ended up with moduleName.moduleName() 

    import costCenterFooter
    footer = costCenterFooter.costCenterFooter(functable)

    for r in headcount_sorted:
        ri = headcount_sorted.index(r)
        #If this is the FIRST row, append the header, then the row
        if ri == 0:
            ws.append(header)
            ws.append(r)
        #If the Cost Center in THIS row is the same as the one before it, append the row
        elif headcount_sorted[ri][1] == headcount_sorted[(ri-1)][1]:
            ws.append(r)
        #If this Cost Center is different than the prior row, it's a new Cost Center
        #insert footer (for the previous Cost Center), one spacer, then write the header and append the row
        else:
            ws.append(footer[str(headcount_sorted[ri-1][1])]) #footer for the previous Cost Center
            ws.append(spacer) #spacer for readability
            ws.append(header)
            ws.append(r)
    ws.append(footer[str(headcount_sorted[ri][1])]) #footer for the FINAL Cost Center
    
    #Once all the Cost Centers are done, add in the grand totals for the Functional Area
    import deptTotal
    DeptTotals = deptTotal.deptTotal(footer, tabname)
    ws.append(spacer) #spacer for readability
    ws.append(DeptTotals[tabname]) #Functional Area Totals & Utilization
                
#commented this out while testing the ws.append() method
#that function seems to work better for my purposes. I need to look at
#using it with the headcount.py module
'''
for c in r:
    ci = headcount_sorted[ri].index(c)
    #print "Cell #, Value:", ci, c
    ws.cell(row = ri, column = ci).value = headcount_sorted[ri][ci]
'''   

#My "Main Loop"; running the data through the two functions
#this is almost DEFINETLY sub-optimal, but it'll have to do for now
##Function 1: functionTable ==> makeSubSeaTable/makeNoSubseaTable
time1 = time.time()
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
time2 = time.time()
print "Creation Time  for all Functional Tables was ", time2-time1, "seconds."

##Function 2: create_tabs
time1 = time.time()
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
time2 = time.time()

#print "Length of all tables is", len(subsTable + estSTable + prjCTable + infoTable + procTable + legaTable + engiTable + humaTable + prjMTable + accoTable + ethiTable + hsesTable + qualTable)
print "Creation Time for ALL tabs was ", time2-time1, "seconds."

##Writing that worksheet to a file
target.save(dest_filename)



#Testing how to get at the values in the Functional Tab tables I'm creating with makeNoSubseaTable
'''
for t in prjCTable:
    ti = prjCTable.index(t)
    for cell in prjCTable[ti]:
        print cell.value, cell.address
        print prjCTable
'''


end = time.time()
dur = end - start
print "Total processing time", dur
print "Don't forget to check: there should be NO 2231 employees in the non-Subsea tabs."
print "Also, double check that there are no NON-2231 employees in the Subsea tab."
