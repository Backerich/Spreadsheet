#!/usr/bin/env python3

import openpyxl

import sys
import os
import re
# import time
# import traceback

def exit():
    # Verlässt das Programm
    return sys.exit()


def clear():
    # Räumt den Terminal Output auf
    os.system("cls" if os.name == "nt" else "clear")


def continue_request(wb, output):
    # Abfrage ob Programm beendet werden soll oder weiter laufen soll
    continue_loop = input(output)
    if (continue_loop == 'n') or (continue_loop == 'exit'): # quit Abfrage
        exit()
    else:
        # wenn nicht 'exit' geht es wieder zum Ausgangsmenü zurück
        menu(wb)

def what_sheet(wb):
    this_sheet = input("Welches Sheet willst du benutzen? \n>  ") 

    if this_sheet == 'exit':
        # quit Abfrage
        exit()
    elif this_sheet == "":
        # wenn nur Enter gedrückt wird wird das aktive Sheet als default genommen
        return wb.get_active_sheet()
    else: 
        try:
            # falls der Inputname existiert wird das ausgewählte Sheet benutzt
            return wb.get_sheet_by_name(this_sheet)
        except KeyError:
                # wenn nicht wird das Terminal aufgeräumt, eine Fehlermeldung ausgegeben und es wird erneut view abgefragt
                clear()
                print("\nDas Sheet {} existiert nicht. Versuchen sie es nochmal.\n".format(this_sheet))
                view(wb) 


def list_sheets(wb):
    # Listet alle Sheets in dem workbook auf
    try:
        # Absicherung durch unerwarteten Bug:
        # ERROR: AttributeError: 'NoneType' object has no attribute 'get_sheet_names'
        sheets = wb.get_sheet_names()
    except AttributeError:
        clear()
        print("Versuch es noch einmal.")
        menu(wb)
    counter = 1

    for string in sheets:
        print("Sheet Nummer {}: {}".format(counter, string))
        # counter zur Orientierung des Endusers
        counter += 1


def get_sheet(wb):
    # Zeigt alle Sheets, fragt ab welches benutzt werden soll und gibt dieses wieder
    list_sheets(wb)
    return what_sheet(wb)


def cell_value(sheet, row_first, column_first):
    # Ermittelt durch Koordinaten den Inhalt der ausgewählten Zelle
    return sheet.cell(row=int(row_first), column=int(column_first)).value


# def grid(string_rows):
#     for row in string_rows:
#         string_of_list = ""
#         # Fügt die Reihen Zahlen hinzu
#         string_of_list += str(string_rows.index(row)) + "| "

#         for i in range(0, len(row)):
#             string_of_list += row[i]
#         print(string_of_list)


# def raw_grid(raw_rows, longest):
#     string_rows = []

#     for rows_in_list in raw_rows:
#         list_temp = []
#         for item in rows_in_list:
#             value_length = longest + 1 - len(str(item))
#             list_temp.append(str(item) + " " * value_length + "|")
#         string_rows.append(list_temp)
#     return string_rows


def grid(values, longest):
    # Vorher raw_grid
    string_rows = []
    for rows_in_list in values:
        list_temp = []
        for item in rows_in_list:
            value_length = longest + 1 - len(str(item))
            list_temp.append(str(item) + " " * value_length + "|")
        string_rows.append(list_temp)

    # Vorher grid
    for row in string_rows:
        string_of_list = ""
        # Fügt die Reihen Zahlen hinzu
        string_of_list += str(string_rows.index(row)) + "| "

        for i in range(0, len(row)):
            string_of_list += row[i]
        print(string_of_list)


def get_values(sheet, row_count, column_count):
    all_raw_rows = []
    longest = 0
    all_raw_rows.append(i for i in  range(1, row_count + 1))

    for row in range(1, row_count + 1):
        innerlist = []

        for column in range(1 , column_count + 1):
            values = cell_value(sheet, row, column)

            if values != None:
                innerlist.append(values)

                if len(values) > longest:
                    longest = len(values)

            else:
                innerlist.append("")
        all_raw_rows.append(innerlist)
    return all_raw_rows, longest


def max_sheet(sheet):
    try:
        return sheet.max_row, sheet.max_column
    except AttributeError:
        print("Dieses Sheet hat keine Spalte oder Zeilen!")


def all(wb):
    # Fragt neu gewünschtes Sheet ab
    sheet = get_sheet(wb)
    clear()

    # Ermittelt Maximale Reihe und Spalte
    max_row = max_sheet(sheet)[0]
    max_column = max_sheet(sheet)[1]

    # Zwei Dimensionale Liste mit den den raw Row Werten
    values = get_values(sheet, max_row, max_column)[0]

    # Ermittelt die längste Zelle des Sheets
    longest = get_values(sheet, max_row, max_column)[1]

    # Verarbeitet die Werte zu einem Grid von Strings
    # string_rows = raw_grid(values, longest)
    # grid(string_rows)

    # test new grid
    grid(values, longest)


def position(sheet):
    position_data = input("Geben sie die Positon ihrer Zelle an wie: 'Reihe, Spalte' (z.B. 2,1).\n>  ")

    if position_data != 'exit':
        clear()

        # RE: Muster zum filtern der Koordinaten 
        line = re.compile(r'''
        ^(?P<row>[\d]+),\s*
        (?P<column>[\d]+)$
        ''', re.X|re.M)

        # RE: Vergleicht das Muster mit dem Input und filtert je Reihe und Spalte heraus 
        # und gibt diese Werte and die 'cell_value' Methode weiter
        for match in line.finditer(position_data):
            return cell_value(sheet, match.group('row'), match.group('column'))
    else:
        exit()


def view(wb):
    # Fragt neu gewünschtes Sheet ab
    sheet = get_sheet(wb)
    clear()

    # Enthält den Inhalt der gewünschten Zelle
    position_data = position(sheet)

    if position_data != None:
        # Falls es einen Inhalt gibt:
        return "Ihre Zelle beinhaltet '{}'.".format(position_data)
    else:
        # Falls es keinen gibt:
        return "Diese Zelle gibt es nicht bitte versuchen sie es nochmal."


def ask_workbook():
    workbook_input = input("Gib deinen Pfad zur Datei an oder zieh ihn rein per drag and drop(es sollte eine Exel datei sein also z.B. .xlsx).\n>  ")
    if workbook_input == 'exit':
        exit()
    else:
        try:
            # Falls es den Path gibt verwendet er nun dieses Workbook
            return openpyxl.load_workbook(workbook_input)
        except FileNotFoundError:
            # Falls nicht wird eine Fehlermeldung ausgegeben und nochmal abgefragt
            clear()
            print("Diser Pfad exsistiert nicht. Versuch es nochmal.")
            ask_workbook()

def menu(wb):
    clear()
    while True:
        # Alle möglichen Functionen des Programmes als Menü
        print("Gib 'exit' ein um das Programm zu verlassen.")
        print("Gib 'view' ein um bestimmt Zeilen zu inspizieren.")
        print("Gib 'edit' ein um dein workbook zu ändern.")
        print("Gib 'all' ein um einen überblick von deinem Spreadsheet zu erhalten.")
        user_input = input(">  ").lower()
        
        if user_input == 'exit':
            exit()
        elif user_input == 'view':
            # Falls Input 'view' Abfrage vom view um einzelnen Zelle Inhalte zu betrachen
            clear()
            print(view(wb))
            # Abrage ob das Programm beendet werden soll
            continue_request(wb, "\nWollen sie weiter machen [Y/n]? \n>  ")
        elif user_input == 'edit':
            # Falls input 'edit' wird abgefragt ob der User das Workbook ändern will bzw. in welches
            clear()
            print("Ändere dein workbook.\n")
            # Ersetzt das aktuelle Workbook mit dem neuen
            wb = ask_workbook()
            clear()
        elif user_input == 'all':
            clear()
            # Zeigt alle Inhalte des Sheets an(eine Übersicht)
            all(wb)
            # Abrage ob das Programm beendet werden soll
            continue_request(wb , "\nWollen sie weiter machen [Y/n]? \n>  ")
        else:
            # Fehlermeldung falls es diesen Menüpunkt nicht gibt und läuft die Schleife nochmal durch
            clear()
            print("Den Befehl '{}'' gibt es nicht. Versuchen sie es noch einmal.\n".format(user_input))
            continue
def main():
    clear()
    # DEBUG: Default Workbbok zum Debugen
    wb = openpyxl.load_workbook('Example/example_two.xlsx')
    # Fragt Workbook ab
    # wb = ask_workbook()
    menu(wb)

if __name__ == "__main__":
    main()

