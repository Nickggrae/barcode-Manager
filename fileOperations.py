#Nicholas Wharton
#Barcode Scanner Item Managment
#file operations to manage scanned barcode data
#3/16/2024

import openpyxl
import pandas


#Generate list of item instances in the current file state
def generateItemList(itemFile):

    #open the AMZN file
    book = openpyxl.load_workbook(itemFile)
    sheet = book.active

    itemList = [] #List of all unique items FNSKU from the itemFile

    #Get the number of unique items in the list
    uniqueCountStr = sheet.cell(row=3, column=1).value
    uniqueCount = int(uniqueCountStr.split()[2])

    i = 0
    while i < uniqueCount:
        #pull the FNSKU from the sheet
        itemList.append(sheet.cell(row=6 + i, column=5).value)
        i += 1

    book.close()

    return itemList



def generateItemInstanceList(itemFile):
    itemList = [] #list each record representing an item instance

    #open the AMZN file
    book = openpyxl.load_workbook(itemFile)
    sheet = book.active

    #Get the number of boxes
    boxCount = str(sheet.cell(row=3, column=13).value)

    #Get the number of unique items in the list
    uniqueCountStr = sheet.cell(row=3, column=1).value
    uniqueCount = int(uniqueCountStr.split()[2])

    #loop through each item instance related to a box
    i = 0
    while i < uniqueCount:
        j = 0
        while j < int(boxCount):
            #get the current cells string value
            currentCell = sheet.cell(row=6 + i, column=13 + j).value
            print(str(currentCell))
            #if the cell has a value add each item instance to the list
            if currentCell is not None:
                k = 0
                while k < int(currentCell):
                    #make the item instance record relating its FNSKU with the Box number and add it to the list
                    itemInstance = ["BOX000000" + boxCount,sheet.cell(row=6 + i, column=5).value]
                    itemList.append(itemInstance)
                    
                    k += 1
            
            j += 1

        i += 1

    return itemList



def updateItemBoxMatrix(itemFile, itemList):
    
    #open the AMZN file
    book = openpyxl.load_workbook(itemFile)
    sheet = book.active

    #Get the number of boxes
    boxCount = str(sheet.cell(row=3, column=13).value)

    #Get the number of unique items in the list
    uniqueCountStr = sheet.cell(row=3, column=1).value
    uniqueCount = int(uniqueCountStr.split()[2])

    #Set every cell in the item-box matrix to None
    i = 0
    while i < uniqueCount:
        j = 0
        while j < int(boxCount):
            sheet.cell(row=6 + i, column=13 + j, value=None)
            sheet.cell(row=6 + i, column=13 + j)
            j += 1
        i += 1


    #loop through each item instance related to a box
    i = 0
    while i < len(itemList):
        j = 0
        #loop to find where which row the item is placed in the file
        while j < uniqueCount:
            #compare the FNSKU of the list record with the sheet record
            if itemList[i][1] == sheet.cell(row=6 + j, column=5).value:
                currentCell = sheet.cell(row=6 + j, column=12 + int(itemList[i][0][9])).value
                print("i: " + str(6 + j) + " j: " + str(12 + int(itemList[i][0][9])))
                if currentCell is None:
                    #set the value to 1
                    sheet.cell(row=6 + i, column=13 + j, value="1")
                else:
                    #increment the existing value by 1
                    sheet.cell(row=6 + i, column=13 + j, value=str(int(currentCell) + 1))
                    
            j += 1
    
        i += 1

    return itemList



print(generateItemList("2024-03-13_02-17-15_pack-group-1_99e4d13a-e977-4009-adb0-5ff8d8457be5.xlsx"))
print(generateItemInstanceList("2024-03-13_02-17-15_pack-group-1_99e4d13a-e977-4009-adb0-5ff8d8457be5.xlsx"))
updateItemBoxMatrix("2024-03-13_02-17-15_pack-group-1_99e4d13a-e977-4009-adb0-5ff8d8457be5.xlsx", [['BOX0000005', 'X003UOK4WD'], ['BOX0000005', 'X003UOK4WD'], ['BOX0000005', 'X003UOK4WD'], ['BOX0000005', 'X003UOK4WD'], ['BOX0000005', 'X003ZME0ZR'], ['BOX0000005', 'X003ZME0ZR'], ['BOX0000005', 'X003ZME0ZR'], ['BOX0000005', 'X003ZME0ZR'], ['BOX0000005', 'X003ZME0ZR'], ['BOX0000005', 'X003ZME0ZR'], ['BOX0000005', 'X003TC53GX'], ['BOX0000005', 'X003TC53GX'], ['BOX0000005', 'X003TC53GX'], ['BOX0000005', 'X003TC53GX'], ['BOX0000005', 'X003TC53GX'], ['BOX0000005', 'X003TC53GX'], ['BOX0000005', 'X003TC53GX']])