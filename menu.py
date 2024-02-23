# Nicholas Wharton
# Secure Password Manager
# Menu Controller
# 2/22/2024

import tkinter as tk
from barcode import scanNewSheet
import openpyxl
import pandas

class SheetMenu(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.selectedRowIndex = tk.StringVar()
        self.filename = tk.StringVar()

    def refresh(self, **kwargs):
        for widget in self.winfo_children():
            widget.destroy()

        filename = kwargs.get("data")
        print("KWARGS WORKS IF THIS SHOWS FILENAME: " + str(filename))
        self.filename.set(filename)

        ds = pandas.read_excel(filename)
        fileRows = ds.shape[0] + 1


        tk.Label(self, text="-----------------------------------------").grid(row=0, column=0)
        
        
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
        tk.Entry(self, textvariable=self.selectedRowIndex).grid(row=i+2, column=1)
        tk.Button(self, text="Delete", command=self.delete).grid(row=i+3, column=0)
        tk.Button(self, text="Quit", command=self.quit).grid(row=i+4, column=0)

    def delete(self):
        selectedRowIndex = self.selectedRowIndex.get()
        
        filename = self.filename.get()
        print("DELETE DELETE DELTET DELETET: " + str(filename))
        book = openpyxl.load_workbook(filename)
        sheet = book.active

        sheet.delete_rows(int(selectedRowIndex))

        book.close()

        self.controller.show_frame(SheetMenu, data=filename)

    def quit(self):
       self.controller.quit()


class StartNewSheet(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

    def refresh(self):
        for widget in self.winfo_children():
            widget.destroy()

        tk.Label(self, text="-----------------------------------------").grid(row=0, column=0)
        tk.Label(self, text="Empty Start New Sheet?").grid(row=1, column=0)
        tk.Label(self, text="-----------------------------------------").grid(row=2, column=0)
        tk.Button(self, text="Start Scan", command=self.startNewSheet).grid(row=3, column=0)
        tk.Button(self, text="Quit", command=self.quit).grid(row=5, column=0)

    def startNewSheet(self):
       outputFilename = scanNewSheet()
       print ("Output Filename: " + outputFilename)
       self.controller.show_frame(SheetMenu, data=outputFilename)

    def quit(self):
       self.controller.quit()




class Menu(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

    def refresh(self):
        for widget in self.winfo_children():
            widget.destroy()

        tk.Label(self, text="-----------------------------------------").grid(row=0, column=0)
        tk.Label(self, text="Would You Like To Do?").grid(row=1, column=0)
        tk.Label(self, text="-----------------------------------------").grid(row=2, column=0)
        tk.Button(self, text="Start a New Sheet", command=self.startNewSheet).grid(row=3, column=0)
        tk.Button(self, text="Quit", command=self.quit).grid(row=5, column=0)

    def startNewSheet(self):
        self.controller.show_frame(StartNewSheet)

    def quit(self):
        self.controller.quit()





class Application(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        self.frames = {}
        for F in (Menu, StartNewSheet, SheetMenu):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(Menu)

    def show_frame(self, cont, **kwargs):
        frame = self.frames[cont]
        frame.refresh(**kwargs)
        frame.tkraise()





if __name__ == "__main__":
    app = Application()
    app.geometry("300x600")
    app.mainloop()
