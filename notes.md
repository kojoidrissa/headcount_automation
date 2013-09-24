##deptTotal Planning Notes
These are the planning notes for this branch (deptTotal)

To create the department (functional area) totals, I need to:

-  Create a seperate function, perhaps a seperate module to handle this
-  Have that module total the footers for each functional area, probably in the 'create_tabs' function. This will give the doeTotal and projTotal for the entire Functional Area
    -  What inputs will this function take? It should be the "footer" dictionary, which should take the form:

            footer = {'cc1':[doeTotal, projTotal}, 'cc2:'[doeTotal, projTotal]...'ccN':[doeTotal, projTotal}

-  append the deptTotal to the end of the tab after the other cost centers are created