#Nicholas Wharton
#Barcode Scanner Item Managment
#Process scanned info
#2/16/2024

import openpyxl
import pandas
from datetime import datetime


def process (scannedItems):
    itemsFile = '2024-02-04_00-29-15_pack-group-1_342d6769-dbe7-43e0-bba3-61cbaeaed180-completed.xlsx' #file that holds the records of what items exist
    boxFile = 'boxes2-16-24.xlsx' #filename of the file that holds all the used box FNSKUs
    boxes = [] #list of all the boxes from the boxFile

    tempCompData = [] #holds a item record with its pulled item data
    compDataArr = [] #list of compiled item records
    currentBox = "" #the current box in use

    print("Starting Processing of Scanned Items...")
    
    #how many lines are there in the boxFile
    ds = pandas.read_excel(boxFile)
    fileRows = ds.shape[0] + 1

    #open the boxfile
    book = openpyxl.load_workbook(boxFile)
    sheet = book.active

    #store all the box names into a list of strings
    i = 1
    while i <= fileRows:
        boxes.append(sheet.cell(row=i, column=1).value)
        i += 1

    book.close()

    #--------------------------------------------------------------------------------------------

    #how many lines are there in the item file
    ds = pandas.read_excel(itemsFile)
    fileRows = ds.shape[0] + 1

    #open the item records file
    book = openpyxl.load_workbook(itemsFile)
    sheet = book.active

    #Process each scanned Item
    for item in scannedItems:
        isBox = False

        #check if the FNSKU is related to a box
        for box in boxes:
            if item == box:
                isBox = True

        #if a box update the current box var and continue to the next scanned item
        if (isBox == True):
            currentBox = item
            continue

        print("items: " + item)

        #search for the item based on its FNSKU
        i = 1 #iterate through each row
        while i <= fileRows:
            if item == sheet.cell(row=i, column=5).value:
                tempCompData.append(currentBox)
                tempCompData.append(sheet.cell(row=i, column=1).value)
                tempCompData.append(sheet.cell(row=i, column=2).value)
                tempCompData.append(sheet.cell(row=i, column=4).value)
                tempCompData.append(sheet.cell(row=i, column=5).value)
            i += 1
        
        #if the item has no data in the record sheet then still add it to the sheet but showing no data was found
        if len(tempCompData) == 0:
            print("Item " + item + " Has No Info In Item Record Field")
            tempCompData.append(currentBox)
            tempCompData.append("N/A")
            tempCompData.append("N/A")
            tempCompData.append("N/A")
            tempCompData.append(item)

        compDataArr.append(tempCompData)
        tempCompData = []
            
    book.close()


    book = openpyxl.Workbook()
    sheet = book.active

    #print the compiled inforamtion and also put it into a sheet
    print("Size of compdataarr: " + str(len(compDataArr)))
    i = 1
    for item in compDataArr:
        j = 1
        for part in item:
            print(part)
            sheet.cell(row=i, column=j, value=part)
            j += 1
        i += 1
        print("\n")

    now = datetime.now()
    outputFilename = 'processed' + str(now.strftime("%m-%d-%H-%M")) + '.xlsx'
    book.save(outputFilename)
    return outputFilename

#test function call to skip barcode.py
#process(["BOX0000001", "X002Z4A4SD", "X002Z8FDVH", "X002YCVC3R", "BOX0000002", "X003UOHSL3", "X003UOK4WD"])