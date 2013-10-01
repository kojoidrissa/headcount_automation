#`formats`  branch planning notes
These are the planning notes for my former "formats" branch

THIS helps a lot: http://stackoverflow.com/questions/8440284/setting-styles-in-openpyxl

1.  DONE!
    2. Experiment with making the Headers bold. I'll make changes to the `#creating the header row` of *functionalSheets.py*, which is lines 13-20 before I start making changes
`dir(cell.style.font)`
...
`'bold',
 'color',
 'italic',
 'name',
 'size',
 'strikethrough',
 'subscript',
 'superscript',
 'underline'`


2. After that, I'll try Text alignment with the headers too
    - Left with Columns 1-6:  `cell.style.alignment.HORIZONTAL_LEFT = 'left'`
    - Right with Columns 7-11: `cell.style.alignment.HORIZONTAL_RIGHT = 'right'`
3.  Number formats (% and Acct) for Util. & Hours
    4. from notes at https://bitbucket.org/ericgazoni/openpyxl/src/c58a6418d68fa1e30d7aa01bd0e0e6cfd14a20ee/openpyxl/style.py?at=default, the following shold work 
        5. %:`cell.style.NumberFormat.FORMAT_PERCENTAGE = '0%'`
        6. Acct: `_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)`
        
My decision to NOT use cell objects (as I didn't know how at the time) may have hamstrung me here. I may need to do the formatting manually.
NOPE! Once I run thing through the `create_tabs` function, they become openpyxl.cell objects. :-)