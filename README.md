#Headcount Summary

This code was initially (2013-Jun) designed for a specific purpose: to summarize detailed time sheet data, then separate that summarized data into specific, predefined categories. As it was created under time-pressure, many shortcuts were taken and many options were hard-coded.

Going forward, I'll be generalizing it so it should work for most types of data.  The primary files here are:

-  [headcountSummary.py](headcountSummary.py): this opens the spreadsheet file, summarizes the data on pre-set criteria (this needs to be generalized), and writes that result to a new spreadsheet file.
-  [functionalSheets.py](functionalSheets.py): this takes the result from headcountSummar.py and sorts it into a number of spreadsheet tabs. That sorting is based on information found in a configuration file, now called `costCenter_Function_map.json`. A few custom modules are called on to do some formatting.  With advances in the OpenPyXL library (as of 2015-01-14), those modules may no longer be needed. The results are then written to a new spreadsheet file.
-  [costCenter_Function_map.json](costCenter_Function_map.json): This contains the map between the sorting sub-categories and the 'super-categories' that they belong to. It also needs to be generalized, including a new name.