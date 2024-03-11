#Nicholas Wharton
#Barcode Scanner Item Managment
#Driver program to recieve scanned info
#3/6/2024

from pynput import keyboard
from processItems import process

scannedItems = [] #list of all of the FNSKUs of the scanned items

currentItem = "" #the temp var holding the most recently scanned items FNSKU
currentItemIndex = 0 #holds the next index to put the next character when buildimg the FNSKU from the scanner input

#exception thrown when '/' is pressed to start the processing of the scanned items
class EscapedExit(Exception):
    pass

#when any key is pressed
def on_press(key):
    global scannedItems
    global currentItemIndex
    global currentItem

    #if the slash key is pressed end the lisening by throwing an exception
    if str(key) == "'/'":
        print("----------Ran Escape--------")
        raise EscapedExit("User Requested Exit")

    #if the key pressed is a number or letter
    if len(str(key)) == 3 and str(key)[1].isalnum():
        #print(str(key))

        #build the FNSKU from scanner inputs or reset and add the value to the list of scanned items
        if currentItemIndex == 9:
            currentItem += str(key)[1]
            currentItemIndex = 0
            scannedItems.append(currentItem)
            currentItem = ""
        else:
            currentItem += str(key)[1]
            currentItemIndex += 1


def scanSheet(itemFile, boxFile, filename=None):

    #-----for testing without scanner------
    #testSet = ["BOX0000001", "X002YCVC3R", "X002Z4A4SD", "BOX0000003", "X003UOHSL3", "BOX0000001", "X002YCVC3R", "BOX0000003", "X002YCVC3R", "BOX0000005", "X003TC53GX", "X003Y3QZ19", "BOX0000007", "X0044EDIRV", "X0044ECG1Z"]
    testSet = ["BOX0000001", "X002YCVC3R", "X002Z4A4SD", "BOX0000004", "X002YCVC3R", "X002Z4A4SD", "BOX0000002", "X003UOHSL3", "BOX0000008", "X002YCVC3R", "BOX0000007", "X002YCVC3R", "BOX0000005", "X003TC53GX", "X003Y3QZ19", "BOX0000009", "X0044EDIRV", "X0044ECG1Z", "BOX0000009", "X0044EDIRV", "X0044ECG1Z"]
    if filename is None:
        outputFilename = process(testSet, itemFile, boxFile)
    else:
        outputFilename = process(testSet, itemFile, boxFile, filename)

    return outputFilename
    #-----for testing without scanner------





    #reset scanned item array for new session
    global scannedItems
    scannedItems = []

    #try to listen for a key unless it is escaped by pressing '/'
    try:
        with keyboard.Listener(on_press=on_press) as listener:
            listener.join()

    #if the user attempts to exit by pressing '/'
    except EscapedExit:

        for item in scannedItems:
            print(item)

        print("Number of Scanned Items: " + str(len(scannedItems)))

        #determine wheter to create a new file or use an existing file when processing the scanned items
        if filename is None:
            outputFilename = process(scannedItems)
        else:
            outputFilename = process(scannedItems, filename)
        
        print("Terminated by User")
        return outputFilename

    #Start lisenting for user input (scanner)
    listener = keyboard.Listener(on_press=on_press)
    listener.start()