#!python
target.remove_sheet('Sheet')
worksheets = target.get_sheet_names()

#Makes Headers and Footers Bold
for sheet in worksheets[:-1]:
	print sheet, len(target.get_sheet_by_name(sheet).rows)
	for row in target.get_sheet_by_name(sheet).rows:
		#ri = target.get_sheet_by_name(sheet).rows.index(row)
		if row[0].value == 'Company': #Header test 
			for cell in row:
				cell.style.font.bold = True
                elif type(row[0].value) == unicode and type(row[5].value) != unicode: #Footer test
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
for sheet in worksheets[:-1]:
	for cell in sheet[0]:
		if cell.value == "DOE":
			di = sheet[0].index(cell)
		elif cell.value == "Project":
			pi = sheet[0].index(cell)
		elif cell.value == "Tot. Hours":
			ti = sheet[0].index(cell)
	for row in target.get_sheet_by_name(sheet).rows:
		if row[5].value + row[6].value == row[7].value:
			print "Match!", row[5].value,"+",row[6].value, "=",row[7].value,"."
			# for cell in row:
			# 	cell.style.font.bold = True
   #              elif type(row[0].value) == unicode and type(row[5].value) != unicode:
		 #    for cell in row[5:]:
		 #        cell.style.font.bold = True

#Results AFTER Workbook has been saved:
'''
for row in target.get_sheet_by_name('Quality').rows[0:3]:
    for cell in row:
        print cell.address, type(cell.value)
        
A1 <type 'unicode'>
B1 <type 'unicode'>
C1 <type 'unicode'>
D1 <type 'unicode'>
E1 <type 'unicode'>
F1 <type 'unicode'>
G1 <type 'unicode'>
H1 <type 'unicode'>
I1 <type 'unicode'>
J1 <type 'unicode'>
A2 <type 'int'>
B2 <type 'int'>
C2 <type 'int'>
D2 <type 'unicode'>
E2 <type 'unicode'>
F2 <type 'int'>
G2 <type 'int'>
H2 <type 'int'>
I2 <type 'float'>
J2 <type 'float'>
A3 <type 'NoneType'>
B3 <type 'NoneType'>
C3 <type 'NoneType'>
D3 <type 'NoneType'>
E3 <type 'NoneType'>
F3 <type 'int'>
G3 <type 'int'>
H3 <type 'float'>
I3 <type 'float'>
J3 <type 'float'>
'''

#Results BEFORE Workbooks has been saved
'''
for row in target.get_sheet_by_name('Quality').rows[0:3]:
    for cell in row:
        print cell.address, type(cell.value)
        
A1 <type 'unicode'>
B1 <type 'unicode'>
C1 <type 'unicode'>
D1 <type 'unicode'>
E1 <type 'unicode'>
F1 <type 'unicode'>
G1 <type 'unicode'>
H1 <type 'unicode'>
I1 <type 'unicode'>
J1 <type 'unicode'>
A2 <type 'int'>
B2 <type 'int'>
C2 <type 'int'>
D2 <type 'unicode'>
E2 <type 'unicode'>
F2 <type 'int'>
G2 <type 'int'>
H2 <type 'int'>
I2 <type 'float'>
J2 <type 'float'>
A3 <type 'unicode'>
B3 <type 'unicode'>
C3 <type 'unicode'>
D3 <type 'unicode'>
E3 <type 'unicode'>
F3 <type 'int'>
G3 <type 'int'>
H3 <type 'float'>
I3 <type 'float'>
J3 <type 'float'>
'''
##So, a lot of the 'unicode' types that are used for blanks BEFORE the save, turn into 'NoneType' aftewards.
##The NoneType later makes sense, as that's what I used in the buffers.