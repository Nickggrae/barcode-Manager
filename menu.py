# Nicholas Wharton
# Secure Password Manager
# Menu Controller
# 3/6/2024

import tkinter as tk
from barcode import scanSheet
import openpyxl
import pandas
import os


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

        tk.Label(self, text="-----------------------------------------").grid(row=0, column=0)
        tk.Label(self, text="Enter Filename: ").grid(row=1, column=0)
        tk.Entry(self, textvariable=self.filename).grid(row=2, column=0)
        tk.Button(self, text="Continue", command=self.submit).grid(row=3, column=0)
        tk.Button(self, text="Quit", command=self.quit).grid(row=4, column=0)
        tk.Label(self, text="-----------------------------------------").grid(row=5, column=0)


    #Start the scanning process on the existing file to append more records
    def submit(self):
        filename = self.filename.get()
        print("FILENAME HERHERHERE: " + filename)
        self.controller.show_frame(SheetMenu, data=filename)

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

    #used when the page is loaded to refresh the displayed information
    def refresh(self, **kwargs):
        #destory the previous window information
        for widget in self.winfo_children():
            widget.destroy()

        filename = ""

        #get the user input if possible
        try:
            filename = kwargs.get("data")
        except:
            filename = self.filename.get()
        
        #determine if the filename was input
        if filename is None:
            filename = self.filename.get()

        print("KWARGS WORKS IF THIS SHOWS FILENAME: " + str(filename))
        self.filename.set(filename)

        #determine how many records are in the file
        ds = pandas.read_excel(filename)
        fileRows = ds.shape[0] + 1


        tk.Label(self, text="-----------------------------------------").grid(row=0, column=0)
        
        #open the book for reading the lines
        book = openpyxl.load_workbook(filename)
        sheet = book.active

        #store all the box names into a list of strings
        i = 1
        while i <= fileRows:
            tk.Label(self, text=str(i) + ": " + sheet.cell(row=i, column=1).value + ", " + sheet.cell(row=i, column=2).value).grid(row=i, column=0)
            i += 1

        book.close()

        tk.Label(self, text="-----------------------------------------").grid(row=i+1, column=0)
        tk.Label(self, text="Select Row: ").grid(row=i+2, column=0)
        tk.Entry(self, textvariable=self.selectedRowIndex).grid(row=i+3, column=0)
        tk.Button(self, text="Delete", command=self.delete).grid(row=i+4, column=0)
        tk.Button(self, text="Add More", command=self.append).grid(row=i+5, column=0)
        tk.Button(self, text="Quit", command=self.quit).grid(row=i+6, column=0)

    #delete a record from the sheet based on user input and refresh the page
    def delete(self):
        selectedRowIndex = self.selectedRowIndex.get()
        filename = self.filename.get()

        #Open the sheet and delete the row based on the users input 
        book = openpyxl.load_workbook(filename)
        sheet = book.active

        sheet.delete_rows(int(selectedRowIndex))

        book.save(filename)
        book.close()

        #update the page info
        self.refresh()

    #Start the scanning process on the existing file to append more records
    def append(self):
        filename = self.filename.get()
        scanSheet(filename)

        self.refresh()

    #end the program
    def quit(self):
        filename = self.filename.get()

        os.system(f'start excel {filename}')

        self.controller.quit()


#Page that lets you start the scanning process for a new sheet
class StartNewSheet(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller #controller of the window

    #refresh the page information 
    def refresh(self):
        #delete the previous window information
        for widget in self.winfo_children():
            widget.destroy()

        tk.Label(self, text="-----------------------------------------").grid(row=0, column=0)
        tk.Label(self, text="Empty Start New Sheet?").grid(row=1, column=0)
        tk.Label(self, text="-----------------------------------------").grid(row=2, column=0)
        tk.Button(self, text="Start Scan", command=self.startNewSheet).grid(row=3, column=0)
        tk.Button(self, text="Quit", command=self.quit).grid(row=5, column=0)

    #start the scanning process to be added to a newly created sheet, then open the sheet menu for the created sheet
    def startNewSheet(self):
       outputFilename = scanSheet()
       print ("Output Filename: " + outputFilename)
       self.controller.show_frame(SheetMenu, data=outputFilename)

    #end the program
    def quit(self):
       self.controller.quit()



#The beginning menu where you can decide to start a new sheet or work off an existing one.
class Menu(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller #controller of the window

    #refresh the information presented on the page
    def refresh(self):
        #delete the previous window information 
        for widget in self.winfo_children():
            widget.destroy()

        tk.Label(self, text="-----------------------------------------").grid(row=0, column=0)
        tk.Label(self, text="Would You Like To Do?").grid(row=1, column=0)
        tk.Label(self, text="-----------------------------------------").grid(row=2, column=0)
        tk.Button(self, text="Start a New Sheet", command=self.startNewSheet).grid(row=3, column=0)
        tk.Button(self, text="Use Existing Sheet", command=self.useExistingSheet).grid(row=4, column=0)

        tk.Button(self, text="Quit", command=self.quit).grid(row=6, column=0)

    #open the start new sheet menu
    def startNewSheet(self):
        self.controller.show_frame(StartNewSheet)

    #open the use an existing sheet menu
    def useExistingSheet(self):
        self.controller.show_frame(GetFilename)


    #end the program
    def quit(self):
        self.controller.quit()


#master class for the application
class Application(tk.Tk):
    #initalize the application window and set the basic setting
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self) #controller of the window
        container.pack(side="top", fill="both", expand=True)

        self.frames = {}
        #initalize which frames exist
        for F in (Menu, StartNewSheet, SheetMenu, GetFilename):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        #the first frame to open on startup
        self.show_frame(Menu)

    #function for switching between frames
    def show_frame(self, cont, **kwargs):
        frame = self.frames[cont]
        frame.refresh(**kwargs)
        frame.tkraise()





if __name__ == "__main__":
    app = Application()
    app.geometry("300x600")
    app.mainloop()
