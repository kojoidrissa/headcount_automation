def summary_not_in_map (fullTable, map):
	'''
	nested list, dictionary --> list

	Takes in the 'fullTable' nested list and the .json map of Cost Centers to Departments
	Compares the cost centers in each to find any cost centers that are IN the summary but are
	NOT in the map

	returns a list of exceptions found
	'''

    #Creating list of unique Cost Centers from the Summary
	summary_ccList = []
	for i in fullTable:
	    summary_ccList.append(unicode(i[1]))
	summaryUnique = set(summary_ccList)

	#Creating list of unique Cost Centers from the map
	map_ccList = []
	for key in dept_dict.keys():
		for i in dept_dict[key]:
			map_ccList.append(i)
	mapUnique = set(map_ccList)

	exception_list = []
	for i in summaryUnique:
		if i not in mapUnique:
			exception_list.append(i)
	exceptions = sorted(exception_list) #http://docs.python.org/2/library/functions.html?highlight=sorted#sorted

	print len(exceptions), "Cost Centers in the Headcount Summary are NOT inlcuded in the Department/Cost Center mapping"
	return exceptions



	#I'm trying to figure out how to build a 'while' loop. Later for that.
	# for i in summaryUnique:
	# 	for key in dept_dict.keys():
	# 		if i not in dept_dict[key]:
	# 			print i, "was not found in", key





