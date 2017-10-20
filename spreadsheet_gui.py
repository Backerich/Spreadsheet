from tkinter import *
from tkinter import ttk
from tkinter import filedialog

# TODO: Replace zeichen selber auswählen können

class App():
    def __init__(self):
        root = Tk()
        root.title("Spreadsheet")
        app = Main(root)
        app.set_up()
        root.mainloop()


class Main():
    def __init__(self, master):
        self.master = master

        self.mainframe = ttk.Frame(master, padding="3 3 12 12")
        self.mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        self.mainframe.columnconfigure(0, weight=1)
        self.mainframe.rowconfigure(0, weight=1)
        self.mainframe.pack()

        self.first_file = StringVar()
        self.second_file = StringVar()

    def file_one(self):
        self.master.filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                   filetypes=(("jpeg files", "*.jpg"), ("all files", "*.*")))
        self.first_file.set(self.master.filename)

    def file_two(self):
        self.master.filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                          filetypes=(("jpeg files", "*.jpg"), ("all files", "*.*")))
        self.second_file.set(self.master.filename)

    def set_up(self):
        label = ttk.Label(self.mainframe, text="Gibt deine Spreadsheets an die du verwenden möchtest...")
        first_workbook = ttk.Entry(self.mainframe, width=60, textvariable=self.first_file)
        second_workbook = ttk.Entry(self.mainframe, width=60, textvariable=self.second_file)
        first_button = ttk.Button(self.mainframe, text="...", command=self.file_one, width=1)
        second_button = ttk.Button(self.mainframe, text="...", command=self.file_two, width=1)
        replace_button = ttk.Button(self.mainframe, text="Replace", )

        label.grid(column=1, row=1)
        first_workbook.grid(column=1, row=2, sticky=(N, W))
        second_workbook.grid(column=1, row=3, sticky=(N, W))
        first_button.grid(column=2, row=2, sticky=(N, W))
        second_button.grid(column=2, row=3, sticky=(N, W))
