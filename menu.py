# Nicholas Wharton
# Secure Password Manager
# Menu Controller
# 3/11/2024

import tkinter as tk
from barcode import scanSheet
from fileOperations import appendNewItem
from fileOperations import deleteRecord
from fileOperations import copyInit
import openpyxl
import pandas
import os

fontSize = 15

#menu for dealing with a sheet that already exists
class GetFilename(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller #controller of the window
        self.filename = tk.StringVar() #the filename of the book holding the scanned records

    #used when the page is loaded to refresh the displayed information
    def refresh(self):
        #destory the previous window information
        for widget in self.winfo_children():
            widget.destroy()

        tk.Label(self, text="Please Enter The Amazon Sheet or Copied Working Sheet", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Enter Filename: ", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Entry(self, textvariable=self.filename, font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Button(self, text="Continue", font=("Helvetica", fontSize), command=self.submit).pack(anchor="center")
        tk.Button(self, text="Quit", font=("Helvetica", fontSize), command=self.quit).pack(anchor="center")
        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")

    #Start the scanning process on the existing file to append more records
    def submit(self):
        filename = self.filename.get()
        
        if filename[0] != 'c' and filename[0] != 'o':
            newFile = copyInit(filename)
            self.controller.show_frame(SheetMenu, processedFilename=newFile)
        else:
            self.controller.show_frame(SheetMenu, processedFilename=filename)

    #end the program
    def quit(self):
       self.controller.quit()



#menu for dealing with a sheet that already exists
class SheetMenu(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller #controller of the window
        self.selectedRowIndex = tk.StringVar() #used to pass the users input to the processing functions
        self.filename = tk.StringVar() #the filename of the book holding the scanned records
        self.refreshes = 0
        self.currentBox = "None"
        self.currentItemList = []

    #used when the page is loaded to refresh the displayed information
    def refresh(self, **kwargs):
        #destory the previous window information
        for widget in self.winfo_children():
            widget.destroy()

        v = tk.Scrollbar(self)
        v.pack(side = "right", fill = "y")
        filename = ""

        #get the user input if possible
        if self.refreshes == 0:
            filename = str(kwargs.get("processedFilename"))
            self.filename.set(filename)
            self.refreshes = 1
        else:
            filename = self.filename.get()

        t = tk.Text(self, width = 15, height = 15, wrap = "none", yscrollcommand = v.set, font=("Helvetica", fontSize))

        t.insert("end", "-----------------------------------------\n")
        
        #open the book for reading the lines
        book = openpyxl.load_workbook(filename)
        sheet = book.active

        #Get the current number of boxes in the sheet.
        currentBoxNum = str(sheet.cell(row=1, column=2).value)

        #Get the number of item records
        ds = pandas.read_excel(filename)
        fileRows = ds.shape[0] + 1
        currentItemNum = fileRows - 2

        #list of current items records pulled from the file
        currentItemList = []

        #Reads in the values from the current working sheet to display the current item instances in the menu
        i = 3
        while (i < (3 + currentItemNum)):
            j = 13
            while (j < (13 + int(currentBoxNum))):
                if sheet.cell(row=i, column=j).value is not None:
                    #get the value of the cell and add the corresponding amount of records
                    z = int(sheet.cell(row=i, column=j).value)
                    while z > 0:
                        currentItemList.append(["BOX000000" + str(j - 12), sheet.cell(row=i, column=5).value])
                        t.insert("end", str(len(currentItemList)) + ": " + "BOX000000" + str(j - 12) + ", " + sheet.cell(row=i, column=5).value + "\n")
                        z -= 1
                j += 1
            
            i += 1

        t.pack(side="left", fill="both", expand=True)
        v.config(command=t.yview)

        book.close()

        #set currentItemList of window object so i can be accessed by the delete function
        self.currentItemList = currentItemList

        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Select Row: ", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Entry(self, textvariable=self.selectedRowIndex, font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Button(self, text="Delete", font=("Helvetica", fontSize), command=self.delete).pack(anchor="center")
        tk.Button(self, text="Add More", font=("Helvetica", fontSize), command=self.append).pack(anchor="center")
        tk.Button(self, text="Quit", font=("Helvetica", fontSize), command=self.quit).pack(anchor="center")
        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Current Box Count: " + currentBoxNum, font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Button(self, text="Add Box", font=("Helvetica", fontSize), command=self.addBox).pack(anchor="center")
        tk.Label(self, text="Current Box:", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text=self.currentBox, font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Now you are editing the created file directly", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Press Add More to enter *Scanning Mode*", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="When in *Scanning Mode* you can scan a ", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="box then all the items for that box", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="You can scan a new box to change the", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="current box and then continue scanning items", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="To leave *Scanning Mode* Press '/'", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="If you need to reenter Scanning Mode just", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="press Add More again.", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="If you click outside the program window", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="while in scanning mode it will become", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="non-responsive just click on any other", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="program and press '/' to leave *Scanning Mode*", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="To delete a record type the index number", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Listed in the input box and press delete", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="To end the program press the quit button", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="This will open the working file for you to see", font=("Helvetica", fontSize)).pack(anchor="center")

    #delete a record from the sheet based on user input and refresh the page
    def delete(self):
        deleteRecord(self.filename.get(), self.selectedRowIndex.get(), self.currentItemList)
        self.refresh()

    #Start the scanning process on the existing file to append more records
    def append(self):
        filename = self.filename.get()
        
        currentBox = ""
        currentItem = ""
        while True:
            item = str(scanSheet()) #get the next scanned item

            if item == "'/'" or item == "": #make sure it was a barcode
                break
            else:
                if item[0] == 'B':
                    currentBox = item
                    self.currentBox = currentBox
                else:
                    currentItem = item
                    if currentBox == "":
                        print("No Box Associated with Item")
                    else:
                        print("currentBox " + currentBox + ", currentItem: " + currentItem)
                        appendNewItem(filename, currentBox, currentItem)
                        currentItem = ""
                self.refresh()
                self.update()
        
        self.currentBox = "None"
        self.refresh()

    #Increments the total boxes value in the working sheet
    def addBox(self):
        filename = self.filename.get()
        
        #open the book for reading the lines
        book = openpyxl.load_workbook(filename)
        sheet = book.active

        sheet.cell(row=1, column=2).value = int(sheet.cell(row=1, column=2).value) + 1
        
        book.save(filename)
        book.close()
        
        self.refresh()

    #end the program
    def quit(self):
        os.system(f'start excel {self.filename.get()}')
        self.controller.quit()


#master class for the application
class Application(tk.Tk):
    #initalize the application window and set the basic setting
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self) #controller of the window
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        #initalize which frames exist
        for F in (SheetMenu, GetFilename):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        #the first frame to open on startup
        self.show_frame(GetFilename)

    #function for switching between frames
    def show_frame(self, cont, **kwargs):
        frame = self.frames[cont]
        frame.refresh(**kwargs)
        frame.tkraise()



if __name__ == "__main__":
    app = Application()
    app.geometry("1000x900")
    app.mainloop()