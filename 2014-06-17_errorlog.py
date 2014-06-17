%run C:/Users/kidrissa/projects/headcount_automation/functionalSheets.py
fullTable Creation Time was  9.61299991608 seconds.
Creation Time  for all Functional Tables was  16.6949999332 seconds.
---------------------------------------------------------------------------
IndexError                                Traceback (most recent call last)
C:\Users\kidrissa\AppData\Local\Enthought\Canopy32\App\appdata\canopy-1.3.0.1715.win-x86\lib\site-packages\IPython\utils\py3compat.pyc in execfile(fname, glob, loc)
    195             else:
    196                 filename = fname
--> 197             exec compile(scripttext, filename, 'exec') in glob, loc
    198     else:
    199         def execfile(fname, *where):

C:\Users\kidrissa\projects\headcount_automation\functionalSheets.py in <module>()
    221 time1 = time.time()
    222 for key in sorted(sheet_dict.keys(), reverse = True):
--> 223     create_tabs(sheet_dict[key], key)
    224 create_tabs(fullTable, 'Headcount Summary Sorted')
    225 time2 = time.time()

C:\Users\kidrissa\projects\headcount_automation\functionalSheets.py in create_tabs(functable, tabname)
    165 
    166     #spacer added to create break for manual insertion of Cost Center sum functions
--> 167     spacer = [None for i in range(len(functable[0]))] #used range in a list comprehension to build this
    168 
    169     #bring in my custom Footer code and generate the Footer dictionary for this Functional area

IndexError: list index out of range 
