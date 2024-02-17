#Nicholas Wharton
#Barcode Scanner Item Managment
#Driver program to recieve scanned info
#2/16/2024

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

#try to listen for a key unless it is escaped by pressing '/'
try:
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

except EscapedExit:
    for item in scannedItems:
        print(item)

    print("Number of Scanned Items: " + str(len(scannedItems)))
    process(scannedItems)
    
    print("Terminated by User")

#Start lisenting for user input (scanner)
listener = keyboard.Listener(on_press=on_press)
listener.start()