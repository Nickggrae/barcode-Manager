# Nicholas Wharton
# Secure Password Manager
# Menu Controller
# 3/6/2024

import tkinter as tk
from barcode import scanSheet
import openpyxl
import pandas
import os
from buildInvoice import buildInvoice
from invoiceToProcessed import invoiceToProcessed

itemsFile = 'SKU-Source-File.xlsx' #file that holds the records of what items exist
boxFile = 'boxes2-16-24.xlsx' #filename of the file that holds all the used box FNSKUs

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

    #used when the page is loaded to refresh the displayed information
    def refresh(self, **kwargs):
        #destory the previous window information
        for widget in self.winfo_children():
            widget.destroy()

        v = tk.Scrollbar(self)
        v.pack(side = "right", fill = "y")

        filename = ""
        #get the user input if possible
        filename = str(kwargs.get("processedFilename"))
        
        t = tk.Text(self, width = 15, height = 15, wrap = "none", yscrollcommand = v.set, font=("Helvetica", fontSize))

        #determine if the filename was input
        if filename == "None":
            filename = self.filename.get()
        else:
            self.filename.set(filename)

        if filename[0] == 'i':
            filename = invoiceToProcessed(filename)
            self.filename.set(filename)

        #determine how many records are in the file
        ds = pandas.read_excel(filename)
        fileRows = ds.shape[0] + 1

        t.insert("end", "-----------------------------------------\n")
        
        #open the book for reading the lines
        book = openpyxl.load_workbook(filename)
        sheet = book.active

        #store all the box names into a list of strings
        i = 1
        while i <= fileRows:
            t.insert("end", str(i) + ": " + sheet.cell(row=i, column=1).value + ", " + sheet.cell(row=i, column=3).value + "\n")
            i += 1

        t.pack(side="left", fill="both", expand=True)
        v.config(command=t.yview)

        book.close()

        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Select Row: ", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Entry(self, textvariable=self.selectedRowIndex, font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Button(self, text="Delete", font=("Helvetica", fontSize), command=self.delete).pack(anchor="center")
        tk.Button(self, text="Add More", font=("Helvetica", fontSize), command=self.append).pack(anchor="center")
        tk.Button(self, text="Quit", font=("Helvetica", fontSize), command=self.quit).pack(anchor="center")

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
        outputFilename = scanSheet(itemsFile, boxFile, filename)
        
        self.refresh()

    #end the program
    def quit(self):
        invoiceFilename = buildInvoice(self.filename.get(), boxFile)
        os.remove(self.filename.get())
        os.system(f'start excel {invoiceFilename}')

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

        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Empty Start New Sheet?", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Button(self, text="Start Scan", font=("Helvetica", fontSize), command=self.startNewSheet).pack(anchor="center")
        tk.Button(self, text="Quit", font=("Helvetica", fontSize), command=self.quit).pack(anchor="center")

    #start the scanning process to be added to a newly created sheet, then open the sheet menu for the created sheet
    def startNewSheet(self):
       outputFilename = scanSheet(itemsFile, boxFile)
       print ("Output Filename: " + outputFilename)
       self.controller.show_frame(SheetMenu, processedFilename=outputFilename)

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

        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="Would You Like To Do?", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Label(self, text="-----------------------------------------", font=("Helvetica", fontSize)).pack(anchor="center")
        tk.Button(self, text="Start a New Sheet", font=("Helvetica", fontSize), command=self.startNewSheet).pack(anchor="center")
        tk.Button(self, text="Use Existing Sheet", font=("Helvetica", fontSize), command=self.useExistingSheet).pack(anchor="center")

        tk.Button(self, text="Quit", font=("Helvetica", fontSize), command=self.quit).pack(anchor="center")

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
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

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
    app.geometry("1000x600")
    app.mainloop()
