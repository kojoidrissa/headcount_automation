##2013-10-18 Errors
`functionalSheets.py` in the Master branch is throwing the following errors when run against hdcntsum.xlsx:

<pre>
%run C:/Users/kidrissa/headcount_automation/functionalSheets.py
---------------------------------------------------------------------------
ZeroDivisionError                         Traceback (most recent call last)
C:\Users\kidrissa\AppData\Local\Enthought\Canopy32\System\lib\site-packages\IPython\utils\py3compat.pyc in execfile(fname, glob, loc)
    195             else:
    196                 filename = fname
--> 197             exec compile(scripttext, filename, 'exec') in glob, loc
    198     else:
    199         def execfile(fname, *where):

C:\Users\kidrissa\headcount_automation\functionalSheets.py in <module>()
     69         #I'm taking a slice of source.rows to ignore that first row for now
     70     temprow.append(temprow[-1]+temprow[-2]) #Total Hours: sum of DOE & Proj Hours
---> 71     temprow.append(temprow[-3]/float(temprow[-1])) #DOE Util%; DOE Hours / newly added Total
     72     temprow.append(temprow[-3]/float(temprow[-2])) #Proj Util%; Proj. Hours / Total
     73     fullTable.append(temprow)

ZeroDivisionError: float division by zero 
</pre>

The same code from the new_stable branch, when run against the identical file, does NOT produce these errors. The error is in the section of code that creates the 'fullTable'. *NOTE: I need to make that cdowe into an actual function*