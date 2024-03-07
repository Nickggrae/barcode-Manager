import tkinter as tk

class Page1(tk.Frame):
    def __init__(self, parent, show_page2):
        tk.Frame.__init__(self, parent)
        tk.Label(self, text="This is Page 1").pack()
        tk.Button(self, text="Go to Page 2", command=show_page2).pack()

class Page2(tk.Frame):
    def __init__(self, parent, show_page1):
        tk.Frame.__init__(self, parent)
        tk.Label(self, text="This is Page 2").pack()
        tk.Button(self, text="Go back to Page 1", command=show_page1).pack()

class MainApp(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("Page Switcher")
        self.geometry("300x200")

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        self.frames = {}
        for F in (Page1, Page2):
            frame = F(container, self.show_page1 if F == Page2 else self.show_page2)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(Page1)

    def show_frame(self, page):
        frame = self.frames[page]
        frame.tkraise()

    def show_page1(self):
        self.show_frame(Page1)

    def show_page2(self):
        self.show_frame(Page2)

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
