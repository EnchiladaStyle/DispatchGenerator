from openpyxl import load_workbook
from travelTimes import TravelTimes
from VRP import VRP

def createDistanceMatrix(ws):

    distanceMatrix = {}
    
    leftList = []
    topList = []
    columns = []
    i = 1
    j = 1
    for cell in ws["A"][1:]:
        leftList.append(cell.value)
        topList.append(cell.value)
        distanceMatrix[cell.value] = {}

    for col in ws.iter_cols(min_row=1, max_row=50, min_col=1, max_col=50):
        columns.append(col)

    while i < len(columns):
        j = 1
        while j < len(columns):
            if columns[i][j].value != None:
                distanceMatrix[leftList[i-1]][topList[j-1]] = columns[i][j].value
            j += 1
        i += 1
    return distanceMatrix


