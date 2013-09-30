#!python
target.remove_sheet('Sheet')
worksheets = target.get_sheet_names()

#Makes Headers and Footers Bold
for sheet in worksheets[:-1]:
	print sheet, len(target.get_sheet_by_name(sheet).rows)
	for row in target.get_sheet_by_name(sheet).rows:
		#ri = target.get_sheet_by_name(sheet).rows.index(row)
		if row[0].value == 'Company': #This is the test for a Header
			for cell in row:
				cell.style.font.bold = True
                elif type(row[0].value) == unicode and type(row[5].value) != unicode:
		    for cell in row[5:]:
		        cell.style.font.bold = True

##Checking the bolding
for sheet in worksheets[:-1]:
	for row in target.get_sheet_by_name(sheet).rows:
		ri = target.get_sheet_by_name(sheet).rows.index(row)
		for cell in row:
			ci = row.index(cell)
			print "Sheet", sheet, "Row", ri, "Cell", ci, ", also known as Cell", cell.address, "is", cell.value
			print "But is it bold?", cell.style.font.bold

#Finding Hours
