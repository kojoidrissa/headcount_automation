def deptTotal(footer, tabname):
    '''
    dict, string --> dict

    Takes in a dictionary (footer) in the following form: 
        footer = {'cc1':[doeTotal1, projTotal1]}...'ccN':[doeTotalN, projTotalN]}
    which is the footer for a Functional area and a string (tabname);
    returns a dict of the following format:
        deptTotal = {'tabname': [sum_doeTotal, sum_projTotal]}

    TEST
        footer = {'cc1' : [1, 2], 'cc2' : [1, 2], 'cc3' : [1, 2], 'cc4' : [1, 2], }
        deptotal = deptTotal(footer, 'Engineering')
        print deptotal
        >>> {'Engineering': [4, 8]}

    '''

    dt_doe = 0
    dt_prj = 0

    for v in footer.itervalues():
        dt_doe = dt_doe + v[-5] #as currently configured, DOETotal is the -5th element in the footer
        dt_prj = dt_prj + v[-4] #as currently configured, ProjTotal is the -4th element in the footer

    deptotal = {tabname: [dt_doe, dt_prj]}
    return deptotal
