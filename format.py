#!python
target.remove_sheet('Sheet')
worksheets = target.get_sheet_names()

for sheet in worksheets[:-1]:
	print sheet, len(target.get_sheet_by_name(sheet).rows)
	for row in target.get_sheet_by_name(sheet).rows:
		ri = target.get_sheet_by_name(sheet).rows.index(row)
		#if row[0].value
		for cell in row:
		    ci = row.index(cell)
		    print "Sheet", sheet, "Row", ri, "Cell", ci, ", also known as Cell", cell.address, "is", cell.value,
		    print "But is it bold?", cell.style.font.bold
	