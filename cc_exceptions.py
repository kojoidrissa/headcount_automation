def summary_not_in_map (fullTable, map):
	'''
	Takes in the 'fullTable' nested list and the .json map of Cost Centers to Departments
	Compares the cost centers in each to find any cost centers that are IN the summary but are
	NOT in the map
	'''

    #Creating list of unique cost Centers
    summary_ccList = []
    for i in fullTable:
        summary_ccList.append(i[1])
    summaryUnique = set(summary_ccList)

    

