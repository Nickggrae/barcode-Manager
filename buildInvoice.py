#Nicholas Wharton
#Barcode Scanner Item Managment
#used to build the invoice from the processed items sheet
#3/7/2024

import openpyxl
import pandas
from datetime import datetime


#driver function for building a invoice from the processed item information
def buildInvoice(inputFile, boxes):
    
    uniqueItems, itemsBoxes = getItemInfo(inputFile) #get the item information needed to build the invoice from the input file

    generateInvoice(uniqueItems, itemsBoxes, inputFile, boxes) #generate the invoice with the collected item information



#Get the item information needed to build the invoice from the input file
def getItemInfo(inputFile):
    uniqueItems = [] #array of unique items that were scanned
    itemsBoxes = [] #array of arrays of each related to the uniqueItems index showing which boxes contain the item

    #how many lines are there in the boxFile
    ds = pandas.read_excel(inputFile)
    fileRows = ds.shape[0] + 1

    #open the boxfile
    book = openpyxl.load_workbook(inputFile)
    sheet = book.active

    #store all the box names into a list of strings
    i = 1
    while i <= fileRows:
        #grab the first items FNSKU and its corresponding box
        currItem = sheet.cell(row=i, column=6).value
        currItemBox = sheet.cell(row=i, column=1).value

        #check if this item already has a record in the unique items array
        isDistinct = True
        j = 0
        while j < len(uniqueItems):
            if currItem == uniqueItems[j]:
                isDistinct = False
            j += 1

        #if the item is the first instance of the item in the file
        if isDistinct == True:
            uniqueItems.append(currItem) #append the new item to the list of unique items
            itemsBoxes.append([]) #start a new box array for the item
            itemsBoxes[len(itemsBoxes) - 1].append(currItemBox) #add the items box number as the first instance in the item's item box array
        else:
            j = 0
            while j < len(uniqueItems): #find where the items record is stored
                print ("curritem: " + currItem + "    unqiqueItem: " + uniqueItems[j])
                if currItem == uniqueItems[j]:
                    itemsBoxes[j].append(currItemBox) #add the items box to its current box record array
                j += 1

        i += 1

    book.close()

    #print the uniqueItems and itemsBoxes array info to console for testing
    i = 0
    for item in uniqueItems:
        print(str(i) + ": " + item)
        i += 1

    i = 0
    for boxes in itemsBoxes:
        print(str(i) + ": ")
        print(boxes)
        i += 1

    return uniqueItems, itemsBoxes



def generateInvoice(uniqueItems, itemsBoxes, inputFile, boxes):
    #how many lines are there in the boxFile
    ds = pandas.read_excel(inputFile)
    fileRows = ds.shape[0] + 1

    #create a new sheet for the invoice
    book = openpyxl.Workbook()
    sheet = book.active

    #add the labels and basic information to the invoice sheet
    #sheet.cell(row=1, column=1, value="Provide the box details for this pack group below. See the instructions sheet if you have more questions")
    sheet.cell(row=2, column=1, value="Pack Group: 1")
    sheet.cell(row=3, column=1, value="Total SKUs: " + str(len(uniqueItems)) + " (" + str(fileRows) + " units)")

    #Add the item inforamtion labels
    sheet.cell(row=5, column=1, value="SKU")
    sheet.cell(row=5, column=2, value="Product Title")
    sheet.cell(row=5, column=3, value="ID")
    sheet.cell(row=5, column=4, value="ASIN")
    sheet.cell(row=5, column=5, value="FNSKU")
    sheet.cell(row=5, column=6, value="Condition")
    sheet.cell(row=5, column=7, value="Prep Type")
    sheet.cell(row=5, column=8, value="Who Preps Unit?")
    sheet.cell(row=5, column=9, value="Who Labels Unit?")
    sheet.cell(row=5, column=10, value="Expected Quantity")
    sheet.cell(row=5, column=11, value="Boxed Quantity")

    #add the label for each box in the box file
    i = 1
    for box in boxes:
        sheet.cell(row=5, column=11+i, value=box)
        i += 1

    sheet.cell(row=4, column=11, value="Total Box Count:")
    sheet.cell(row=4, column=12, value=str(len(boxes)))

    #open the inputFile to get each items information
    inBook = openpyxl.load_workbook(inputFile)
    inSheet = inBook.active

    #add the information for each unique item into the sheet
    i = 0
    while i < len(uniqueItems):
        j = 1
        #for each item in the input file
        while j <= fileRows:
            #if the item is one of the unique items copy its data into the invoice 
            if uniqueItems[i] == inSheet.cell(row=j, column=6).value:
                sheet.cell(row=i + 6, column=1, value=inSheet.cell(row=j, column=2).value)
                sheet.cell(row=i + 6, column=2, value=inSheet.cell(row=j, column=3).value)
                sheet.cell(row=i + 6, column=3, value=inSheet.cell(row=j, column=4).value)
                sheet.cell(row=i + 6, column=4, value=inSheet.cell(row=j, column=5).value)
                sheet.cell(row=i + 6, column=5, value=inSheet.cell(row=j, column=6).value)
                sheet.cell(row=i + 6, column=6, value=inSheet.cell(row=j, column=7).value)
                sheet.cell(row=i + 6, column=7, value=inSheet.cell(row=j, column=8).value)
                sheet.cell(row=i + 6, column=8, value=inSheet.cell(row=j, column=9).value)
                sheet.cell(row=i + 6, column=9, value=inSheet.cell(row=j, column=10).value)
                sheet.cell(row=i + 6, column=10, value=inSheet.cell(row=j, column=11).value)
                sheet.cell(row=i + 6, column=11, value=inSheet.cell(row=j, column=11).value)
                break
            j += 1

        #fill in the box item association matrix
        for box in itemsBoxes[i]:
            characters = [char for char in box] #break the box string into an array of characters
            boxNumber = int(characters[9]) #get the number value from the box string

            #determine if a value exists for the relation between the item's instance # and the box
            if sheet.cell(row=i + 6, column=boxNumber + 11).value is None:
                #initalize the instance # assocaited with the box to 1
                sheet.cell(row=i + 6, column=boxNumber + 11, value="1")
            else:
                #increment the current instance # assocaited with the box by 1
                sheet.cell(row=i + 6, column=boxNumber + 11, value=str(int(sheet.cell(row=i + 6, column=boxNumber + 11).value) + 1))
                


        i += 1






    #set the colum width for each cell to allow the user to see all of the information when they open the sheet
    columnMaxLen = {}
    for row in sheet.iter_rows(values_only=True):
        for i, cellText in enumerate(row, start=1):
            if i not in columnMaxLen:
                columnMaxLen[i] = len(str(cellText))
            else:
                columnMaxLen[i] = max(columnMaxLen[i], len(str(cellText)))

    # Set column width based on the maximum length of text
    for column, maxLen in columnMaxLen.items():
        sheet.column_dimensions[sheet.cell(row=1, column=column).column_letter].width = maxLen

    


    inBook.close()
    #sheet.cell(row=i, column=j, value=part)


    #generate a name for the invoice and save the file
    now = datetime.now()
    outputFilename = 'invoice' + str(now.strftime("%m-%d-%H-%M")) + '.xlsx'
    book.save(outputFilename)