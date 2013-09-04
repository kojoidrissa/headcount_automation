def costCenterFooter(functable):
    '''
    list of lists --> list of lists

    Takes in a functional area table (functable), finds it's unique set of Cost Centers,
    then uses that list to create a list of key:value pairs in the format {'costCenter': [sumDOE, sumProj]}
    That list is appended as the final element of the functable, which is returned
    '''

    #Creating list of unique cost Centers
    ccList = []
    for i in functable:
        ccList.append(i[1])
    ccUnique = set(ccList)
    
    footer = []
    for cc in ccUnique:
        ccDict = {str(cc):[]}
        sumDoe = 0
        sumProj = 0
        for row in functable:
            if row[1] == cc:
                sumDoe = sumDoe + row[-5]
                sumProj = sumProj + row[-4]
        ccDict[str(cc)].extend([sumDoe, sumProj])
        footer.append(ccDict)
    functable.append(footer)
    return functable


    