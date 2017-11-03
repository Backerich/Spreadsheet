from spreadsheet_logic import Workbook, Sheets

from tkinter import *
from tkinter import ttk
from tkinter import filedialog

# TODO: Replace zeichen selber auswählen können
# TODO: Workbook abfragen und umsetzen
# TODO: file function aufräumen
# TODO: Sheet abfragen -> mit Combobox


class App(object):
    def __init__(self):
        root = Tk()
        root.title("Spreadsheet")
        app = Main(root)
        app.set_up()
        root.mainloop()


class Main(object):
    def __init__(self, master):
        self.master = master

        self.mainframe = ttk.Frame(master, padding="3 3 12 12")
        self.mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        self.mainframe.columnconfigure(0, weight=1)
        self.mainframe.rowconfigure(0, weight=1)
        self.mainframe.pack()

        self.workbook = Workbook()
        self.sheets = Sheets()

        self.first_file = StringVar()
        self.second_file = StringVar()

        self.wb_one = None
        self.wb_two = None

        self.clicked_first = False
        self.clicked_second = False

    def refractor_second(self):
        first_row = 1
        if self.clicked_second:
            self.second_workbook.grid(column=1, row=4, sticky=(N, W))
            self.second_button.grid(column=2, row=4, sticky=(N, W))
            self.second_combobox.grid(column=1, row=5, sticky=(N, W))
            self.replace_button.grid(column=3, row=5, sticky=(N, W))

    def file(self, file, wb, second_clicked, scale=False):
        self.master.filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                          filetypes=(("openoffice files", "*.ods"),
                                                                     ("excel files", "*.xlsx"),
                                                                     ("all files", "*.*")))
        file.set(self.master.filename)

        try:
            wb = self.workbook.get_workbook(file.get())
        except FileNotFoundError:
            print("No")

        combo = ttk.Combobox(self.mainframe, state="readonly", values=self.sheets.list_of_sheets(wb))

        if scale and not(second_clicked):  # Wenn scale und erstes angeklickt
            self.second_workbook.grid(column=1, row=4, sticky=(N, W))  # scale down different
            self.second_button.grid(column=2, row=4, sticky=(N, W))  # scale down different
            self.replace_button.grid(column=3, row=4, sticky=(N, W))  # scale down different

            combo.grid(column=1, row=3, sticky=(N, W))

        elif not(scale) and not(second_clicked):
            combo.grid(column=1, row=5, sticky=(N, W)) # dynamisch machen

        elif not(scale) and second_clicked:
            combo.grid(column=1, row=4, sticky=(N, W))

    def file_one(self):
        self.master.filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                          filetypes=(("openoffice files", "*.ods"),
                                                                     ("excel files", "*.xlsx"),
                                                                     ("all files", "*.*")))
        self.first_file.set(self.master.filename) # first_file different

        try:
            self.wb_one = self.workbook.get_workbook(str(self.first_file.get())) # first_file different, wb different
        except (FileNotFoundError, AttributeError):
            print("No")

        self.second_workbook.grid(column=1, row=4, sticky=(N, W)) # scale down different
        self.second_button.grid(column=2, row=4, sticky=(N, W)) # scale down different
        self.replace_button.grid(column=3, row=4, sticky=(N, W)) # scale down different

        self.first_combobox = ttk.Combobox(self.mainframe, state="readonly", values=
                     self.sheets.list_of_sheets(self.wb_one)).grid(column=1, row=3, sticky=(N, W))  # wb different # position diffierent

        self.refractor_second()

        self.clicked_first = True

    def file_two(self):
        self.master.filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                          filetypes=(("openoffice files", "*.ods"),
                                                                     ("excel files", "*.xlsx"),
                                                                     ("all files", "*.*")))
        self.second_file.set(self.master.filename)

        try:
            self.wb_two = self.workbook.get_workbook(str(self.second_file.get()))
        except (FileNotFoundError, AttributeError):
            print("No")

        row = 4

        if self.clicked_first:
            row = 5

        self.second_combobox = ttk.Combobox(self.mainframe, state="readonly", values=self.sheets.list_of_sheets(self.wb_two))
        self.second_combobox.grid(column=1, row=row, sticky=(N, W))  # wb different # position diffierent

        self.replace_button.grid(column=4, row=row, stick=(N, W))
        self.clicked_second = True

    def set_up(self):
        self.label = ttk.Label(self.mainframe, text="Gibt deine Spreadsheets an die du verwenden möchtest...")
        self.first_workbook = ttk.Entry(self.mainframe, width=60, textvariable=self.first_file)
        self.second_workbook = ttk.Entry(self.mainframe, width=60, textvariable=self.second_file)
        self.first_button = ttk.Button(self.mainframe, text="...", command=self.file_one, width=1)
        self.second_button = ttk.Button(self.mainframe, text="...", command=self.file_two, width=1)
        self.replace_button = ttk.Button(self.mainframe, text="Replace", command=self.replace)

        self.label.grid(column=1, row=1)
        self.first_workbook.grid(column=1, row=2, sticky=(N, W))
        self.second_workbook.grid(column=1, row=3, sticky=(N, W))
        self.first_button.grid(column=2, row=2, sticky=(N, W))
        self.second_button.grid(column=2, row=3, sticky=(N, W))
        self.replace_button.grid(column=3, row=3, sticky=(N, W))

    def replace(self):
        pass

    def set_up_workbook(self):
        sheet_string = ""
        sheet = self.sheets.get_sheet(self.wb_one, sheet_string)
        print(sheet)
