#Nicholas Wharton
#Barcode Scanner Item Managment
#Process scanned info
#3/11/2024

import openpyxl
import pandas
from datetime import datetime


def process (scannedItems, itemsFile, boxFile, inputFile=None):
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
            if currentBox != "": #if a box has not been scanned drop the item record
                if item == sheet.cell(row=i, column=2).value:
                    tempCompData.append(currentBox)
                    tempCompData.append(sheet.cell(row=i, column=1).value)
                    tempCompData.append(sheet.cell(row=i, column=2).value)
                    tempCompData.append(sheet.cell(row=i, column=3).value)
                    tempCompData.append(sheet.cell(row=i, column=4).value)
            i += 1
        
        #if the item has no data in the record sheet then still add it to the sheet but showing no data was found
        if currentBox != "": #if a box has not been scanned drop the item record
            if len(tempCompData) == 0:
                print("Item " + item + " Has No Info In Item Record Field")
                tempCompData.append(currentBox)
                tempCompData.append("N/A")
                tempCompData.append(item)

            compDataArr.append(tempCompData)
        
        tempCompData = []
            
    book.close()

    retVal = ""

    #if the filename was not input create a new file and add the scanned records
    if inputFile == None:
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
        book.close()

        retVal = outputFilename
    
    else:
        ds = pandas.read_excel(inputFile)
        fileRows = ds.shape[0] + 1

        #if a input file is passed append the scanned records to the existing file
        book = openpyxl.load_workbook(inputFile)
        sheet = book.active

        print("Size of compdataarr: " + str(len(compDataArr)))
        i = fileRows + 1
        for item in compDataArr:
            j = 1
            for part in item:
                print(part)
                sheet.cell(row=i, column=j, value=part)
                j += 1
            i += 1
            print("\n")

        book.save(inputFile)
        book.close()

        retVal = inputFile

    return retVal