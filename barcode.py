#Nicholas Wharton
#Barcode Scanner Item Managment
#Driver program to recieve scanned info
#3/11/2024

from pynput import keyboard


currentItem = "" #the temp var holding the most recently scanned items FNSKU
currentItemIndex = 0 #holds the next index to put the next character when buildimg the FNSKU from the scanner input

#exception thrown when '/' is pressed to start the processing of the scanned items
class EscapedExit(Exception):
    pass

#when any key is pressed
def on_press(key):
    global currentItemIndex
    global currentItem

    if str(key) == "'/'":
        print("----------Ran Escape--------")
        currentItem = str(key)
        currentItemIndex = 0
        currentItem = ""
        raise EscapedExit("User Requested Exit")

    #if the key pressed is a number or letter
    if len(str(key)) == 3 and str(key)[1].isalnum():

        print("scanned char: " + str(key)[1])
        #build the FNSKU from scanner inputs or reset and add the value to the list of scanned items
        if currentItemIndex == 9:
            currentItem += str(key)[1].capitalize()
            currentItemIndex = 0
            print("Scanned Item: " + currentItem)
            raise EscapedExit("Barcode Scanned")
        else:
            currentItem += str(key)[1].capitalize()
            currentItemIndex += 1


def scanSheet():
    global currentItem

    #try to listen for a key unless it is escaped by pressing '/'
    try:
        with keyboard.Listener(on_press=on_press) as listener:
            listener.join()

    #if the user attempts to exit by pressing '/'
    except EscapedExit:
        retVal = currentItem
        currentItem = ""
        print("RET VALUE AT ESC: " + retVal)
        return retVal