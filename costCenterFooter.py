def costCenterFooter(functable):
    '''
    list of lists --> dict

    Takes in a functional area table (functable), finds it's unique set of Cost Centers,
    then uses that list to create a list of key:value pairs in the format {'costCenter': [sumDOE, sumProj]}
    Here in the 'footer' branch, I'm trying to return JUST the footer dictionary, not appending it to the table
    '''

    #Creating list of unique cost Centers
    ccList = []
    for i in functable:
        ccList.append(i[1])
    ccUnique = set(ccList)
    
    ##Here in the 'footer' branch, I'm trying to return JUST the footer dictionary, not appending it to the table
    ##The goal is to end up with {'cc1':[doeTotal, projTotal}, 'cc2:'[doeTotal, projTotal]...} and return it
    ##to a new "functFooter" variable, which would be called into the 'create_tabs' function and used to 
    ##calculate the footer row for each cost center
    footer = {}
    for cc in ccUnique:
        ccDict = {str(cc):[]}
        sumDoe = 0
        sumProj = 0
        for row in functable:
            if row[1] == cc:
                sumDoe = sumDoe + row[-5]
                sumProj = sumProj + row[-4]
        ccDict[str(cc)].extend([sumDoe, sumProj]) #StackOverflow helped with this: http://stackoverflow.com/a/3419217 
        footer.append(ccDict)
    return footer


    