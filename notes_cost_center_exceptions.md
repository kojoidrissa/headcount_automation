#notes_cost_center_exceptions.md
Planning notes for the cost_center_exceptions branch

I want to know the following things after the report has been generated:

-  Which cost centers in my .json map are NOT present in the headcount summary? 
    -  This will tell me which cost centers don't have any activity. It may also detect flaws in the summary (something being left out). However, I'm depending more on comparing the total hours between the 'raw data' file and the headcount summary for this.
-  Which cost centers are IN the summary, but NOT in my .json map? **DONE**
    -  This may point out cost centers that aren't being included in their proper department

For some reason, my format.py code isn't picking  up the footers any more. But, THAT bit of the code was always precarious. I'll have to look at it later.
