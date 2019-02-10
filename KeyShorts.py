import tkinter
from tkinter import messagebox
import webbrowser

import xlrd

location = "data\shorts.xlsx"
wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)


class Keyshorts(tkinter.Tk):
    def __init__(self, parent):
        tkinter.Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()
        self.search_query = None
        self.search = None
        self.menubar = None
        self.drop = None
        self.searchbtn = None
        self.searchbox = None
        self.list = None
        self.scrollbar = None

    def initialize(self):
        def update():
            messagebox.showerror("No Update", "This is the alpha version")

        def feedback():
            webbrowser.open('mailto:ristikmajumdar@protonmail.com?Subject=Feedback-To-KeyShorts', new=1)

        def bug():
            webbrowser.open('https://github.com/1bl4z3r/KeyShorts', new=1)

        def how():
            messagebox.showinfo("HowTo", "This is Alpha Version, so no HowTo")

        def about():
            t = tkinter.Toplevel(self)
            t.title("About")
            t.wm_iconbitmap('data\logo.ico')
            t.geometry('300x100')

            more = tkinter.Label(t,
                                 text=u"KeyShorts is a small freeware to provide all possible Keyboard shortcuts of various softwares and Operating Systems. Currently supporting Windows only.",
                                 wraplength=300)
            more.pack()
            dev = tkinter.Label(t, text=u"DEVELOPER-Ristik Majumdar", wraplength=300)
            dev.pack()

            button = tkinter.Button(t, text="Got It!", command=t.destroy)
            button.pack()

        self.grid()

        self.menubar = tkinter.Menu(self.parent)
        self.drop = tkinter.Menu(self.menubar, tearoff=0)
        self.drop.add_command(label="How To Use", command=how)
        self.drop.add_command(label="About", command=about)
        self.drop.add_separator()
        self.drop.add_command(label="Submit Feedback", command=feedback)
        self.drop.add_command(label="Report A Bug", command=bug)
        self.drop.add_command(label="Check For Updates", command=update)

        self.menubar.add_cascade(label="Help", menu=self.drop)
        self.menubar.add_command(label="Exit", command=self.quit)
        self.config(menu=self.menubar)

        self.search_query = tkinter.StringVar()

        self.search = tkinter.Label(self, text=u"Search: ")
        self.search.grid(column=0, row=0, sticky='EW')

        self.searchbox = tkinter.Entry(self, textvariable=self.search_query)
        self.searchbox.grid(column=1, row=0, sticky='EW')
        self.search_query.set(u"Currently Unavailable")

        self.searchbtn = tkinter.Button(self, text=u"Find")  #\U0001F50D
        self.searchbtn.grid(column=2, row=0, sticky='EW')

        self.list = tkinter.Listbox(self)
        self.scrollbar = tkinter.Scrollbar(self, orient="vertical")

        self.list.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.list.yview)

        for i in range(sheet.nrows):
            self.list.insert(tkinter.END, "  ".join(sheet.row_values(i)))

        self.list.grid(row=1, column=0, sticky='NSEW', columnspan=2)
        self.list.columnconfigure(0, weight=1)

        self.scrollbar.grid(column=2, row=1, sticky='NS')
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.resizable(False, True)
        self.update()
        self.geometry(self.geometry())


if __name__ == "__main__":
    app = Keyshorts(None)
    app.title('KeyShorts (0.0.2)')
    app.wm_iconbitmap('data\logo.ico')
    app.geometry('1000x500')
    app.mainloop()
