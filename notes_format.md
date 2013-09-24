#formats Planning Notes
These are the planning notes for my former "formats" branch

THIS helps a lot: http://stackoverflow.com/questions/8440284/setting-styles-in-openpyxl

1.  Experiment with making the Headers bold. I'll make changes to the `#creating the header row` of *functionalSheets.py*, which is lines 13-20 before I start making changes
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
    - Left with Columns 1-6; Right with Columns 7-11
3.  Number formats (% and Acct) for Util. & Hours

My decision to NOT use cell objects (as I didn't know how at the time) may have hamstrung me here. I may need to do the formatting manually.