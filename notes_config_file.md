#`config_file`  branch planning notes

If I'm going to use `costCenter_Function_map.json` as my config file, that'll require some changes to `functionalSheets.py`. I should be able to replace the two big blocks running functions with calls that reference the .json dictionary. But HOW exactly will I do this?

I should be able to replace the:
<pre>
    subsTable = makeSubseaTable(subsea)
    estSTable = makeNoSubseaTable(estSales)
    prjCTable = makeNoSubseaTable(prjCntrls)
    infoTable = makeNoSubseaTable(infotech)
    procTable = makeNoSubseaTable(procure)
    legaTable = makeNoSubseaTable(legal)
    engiTable = makeNoSubseaTable(engineering)
    humaTable = makeNoSubseaTable(humanres)
    prjMTable = makeNoSubseaTable(prjMgmt)
    accoTable = makeNoSubseaTable(accounting)
    ethiTable = makeNoSubseaTable(ethics)
    hsesTable = makeNoSubseaTable(hses)
    qualTable = makeNoSubseaTable(quality)
</pre>