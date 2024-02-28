#insert this into the field = block above expression in arc
#isDuplicate(!FIELDNEEDEDTOFINDDUPE!)

uniquelist = []

def isDuplicate(inValue):
    if inValue in uniquelist:
        return "Duplicated"
    else:
        uniquelist.append(inValue)
        return "GoodData"