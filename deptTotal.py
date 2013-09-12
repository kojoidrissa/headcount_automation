def deptTotal(footer, tabname):
    '''
    dict, string --> dict

    Takes in a dictionary (footer) in the following form: 
        footer = {'cc1':[doeTotal1, projTotal1]}...'ccN':[doeTotalN, projTotalN]}
    which is the footer for a Functional area and a string (tabname);
    returns a dict of the following format:
        deptTotal = {'tabname': [sum_doeTotal, sum_projTotal]}
    '''

    dt_doe = 0
    dt_prj = 0

    for v in footer.itervalues():
        dt_doe = dt_doe + v[0]
        dt_prj = dt_prj + v[1]
