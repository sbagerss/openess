
##################################
def cellPositionFromString(posString):

    if len(posString) < 2:
        raise Exception("Wrong cell position: " + posString)

    if not(ord(posString[0]) >= 65 and ord(posString[0]) <= 90):
        raise Exception("Wrong column number: " + posString[0])

    if not posString[1].isnumeric():
        raise Exception("Wrong row number: " + posString[1])

    row = int(posString[1])
    column = ord(posString[0]) - 64

    return [row,column]

##################################
def cellPositionFromNumbers(pos):
    return chr(pos[1] + 64) + str(pos[0])

