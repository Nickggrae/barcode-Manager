# Nicholas Wharton
# Secure Password Manager
# Menu Controller
# 3/11/2024

import tkinter as tk
from barcode import scanSheet
import openpyxl
import pandas
import os
from datetime import datetime

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

        tk.Label(self, text="Please Enter The Amazon Sheet", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Enter Filename: ", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Entry(self, textvariable=self.filename, font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Button(self, text="Continue", font=("Helvetica", fontSize), command=self.submit).pack(anchor="center")
        tk.Button(self, text="Quit", font=("Helvetica", fontSize), command=self.quit).pack(anchor="center")
        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")


    #Start the scanning process on the existing file to append more records
    def submit(self):
        filename = self.filename.get()
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

        #determine how many records are in the file
        ds = pandas.read_excel(filename)
        fileRows = ds.shape[0] + 1

        t.insert("end", "-----------------------------------------\n")
        
        #open the book for reading the lines
        book = openpyxl.load_workbook(filename)
        sheet = book.active

        #Get the current number of boxes in the sheet.
        currentBoxNum = str(sheet.cell(row=4, column=1).value)

        #Get the number of item records
        currentItemNum = str(sheet.cell(row=4, column=1).value)



        i = 3
        while (i < (3 + currentItemNum)):
            j = 13
            while (j < (13 + currentBoxNum)):
                
                j += 1
            
            i += 1


        #add the items from the processed file to the application frame
        # print("Filerows: " + str(fileRows))
        # i = 1
        # while (i+5) < (fileRows-7):
        #     print("row: " + str(i+5) + ",     column: 1 or 3")
        #     t.insert("end", str(i) + ": " + sheet.cell(row=i+5, column=1).value + ", " + sheet.cell(row=i+5, column=3).value + "\n")
        #     i += 1

        t.pack(side="left", fill="both", expand=True)
        v.config(command=t.yview)

        book.close()

        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Select Row: ", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Entry(self, textvariable=self.selectedRowIndex, font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Button(self, text="Delete", font=("Helvetica", fontSize), command=self.delete).pack(anchor="center")
        tk.Button(self, text="Add More", font=("Helvetica", fontSize), command=self.append).pack(anchor="center")
        tk.Button(self, text="Quit", font=("Helvetica", fontSize), command=self.quit).pack(anchor="center")
        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Current Box:", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text=self.currentBox, font=("Helvetica", fontSize)).pack(anchor="center")

    #delete a record from the sheet based on user input and refresh the page
    def delete(self):
        selectedRowIndex = self.selectedRowIndex.get()
        filename = self.filename.get()

        print("delete filename: " + filename)
        print("row number: " + selectedRowIndex)
        #Open the sheet and delete the row based on the users input 
        book = openpyxl.load_workbook(filename)
        sheet = book.active

        sheet.delete_rows(int(selectedRowIndex) + 1)

        book.save(filename)
        book.close()

        #update the page info
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
                        process(filename, currentBox, currentItem)
                        currentItem = ""
                self.refresh()
                self.update()
        
        self.currentBox = "None"
        self.refresh()

    #end the program
    def quit(self):
        invoiceFilename = buildInvoice(self.filename.get(), boxFile)
        
        os.remove(self.filename.get())
        os.system(f'start excel {invoiceFilename}')

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
    app.geometry("1000x600")
    app.mainloop()