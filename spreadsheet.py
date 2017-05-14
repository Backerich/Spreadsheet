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


def ask_workbook(input_string):
    workbook_input = input(input_string).lower() # "Gib deinen Pfad zur Datei an oder zieh ihn rein per drag and drop(es sollte eine Exel datei sein also z.B. .xlsx).\n>  "
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
            ask_workbook("Gib deinen Pfad zur Datei an oder zieh ihn rein per drag and drop(es sollte eine Exel datei sein also z.B. .xlsx).\n>  ")


def continue_request(wb, output):
    # Abfrage ob Programm beendet werden soll oder weiter laufen soll
    continue_loop = input(output).lower()
    if (continue_loop == 'n') or (continue_loop == 'exit'): # quit Abfrage
        exit()
    else:
        # wenn nicht 'exit' geht es wieder zum Ausgangsmenü zurück
        menu(wb)


def what_sheet(wb):
    this_sheet = input("Welches Sheet willst du benutzen? \n>  ").lower()

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
                get_sheet(wb) 


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


def max_sheet(sheet):
    try:
        # Falls es eine Reihe und Spalte gibt wird die höchste wiedergegeben
        return sheet.max_row, sheet.max_column
    except AttributeError:
        # Falls es keine Reihe oder Spalte gibt wird eine Fahlermeldung ausgegeben
        print("Dieses Sheet hat keine Spalte oder Zeilen!")


def get_sheet(wb):
    # Zeigt alle Sheets, fragt ab welches benutzt werden soll und gibt dieses wieder
    list_sheets(wb)
    return what_sheet(wb)


def cell_value(sheet, row_first, column_first):
    # Ermittelt durch Koordinaten den Inhalt der ausgewählten Zelle
    return sheet.cell(row=int(row_first), column=int(column_first))


def get_values(sheet, row_count, column_count):
    all_raw_rows = []
    longest = 0
    all_raw_rows.append(i for i in  range(1, row_count + 1))

    for row in range(1, row_count + 1):
        innerlist = []

        for column in range(1 , column_count + 1):
            values = cell_value(sheet, row, column).value

            if values != None:
                innerlist.append(values)

                if len(values) > longest:
                    longest = len(values)

            else:
                innerlist.append("")
        all_raw_rows.append(innerlist)
    return all_raw_rows, longest


def copy():
    print("Es wird ihnen eine Ausgangskopie erstellt.")
    copy_name = input("Wie soll ihre Ausgangskopie heißen? (Excelendung erforderlich z.B. xlsx)\n>  ").lower()
    # path = first_workbook_name + " " + copy_name
    # os.system("copy " + path if os.name == "nt" else "cp " + path)
    return "Example/" + copy_name


def compare_sheets(first_values, second_values, first_data, first_workbook):
    first_sheet_values = []
    for first_row in first_values:
        for first_item in first_row:
            if first_item != '': # Wegen Dict kann velue nur ein wert haben
                # So angeben row/column z.B 1,A = 1,1
                position_row = first_values.index(first_row)
                position_column = first_row.index(first_item)
                first_row[position_column] += "done"

                first_sheet_values.append([first_item, (position_row + 1, position_column + 1)])

    for second_row in second_values:
        for second_item in second_row:
            if second_item == '':
                pass
            else:
                line = re.compile(r'''
                    ^\s*(?P<basic_item>[\w\d]+)\s*=\s*
                    (?P<to_change>[\w\d]+)\s*$
                    ''', re.X|re.M)

                for match in line.finditer(second_item):
                    basic = match.group('basic_item')
                    change = match.group('to_change')

                    for cell in first_sheet_values:
                            item = cell[0]
                            if basic == item:
                                row = cell[1][0]
                                column = cell[1][1]
                                pos = cell_value(first_data, row, column).coordinate
                                first_data[pos] = change
                                print(pos)
                                print(first_data[pos].value)

    third_name = copy()
    first_workbook.save(third_name)


# def compare_sheets(first_values, second_values, first_data, first_workbook):
#     first_sheet_values = {}
#     print(first_values)
#     for first_row in first_values:
#         for first_item in first_row:
#             if first_item != '': # Wegen Dict kann key nur ein wert haben key und value tauschen
#                 # So angeben row/column z.B 1,A = 1,1
#                 position_row = first_values.index(first_row) + 1
#                 position_column = first_row.index(first_item) + 1

#                 first_sheet_values[first_item] = (position_row, position_column)

#     for second_row in second_values:
#         for second_item in second_row:
#             if second_item == '':
#                 pass
#             else:
#                 line = re.compile(r'''
#                     ^\s*(?P<basic_item>[\w\d]+)\s*=\s*
#                     (?P<to_change>[\w\d]+)\s*$
#                     ''', re.X|re.M)

#                 for match in line.finditer(second_item):
#                     basic = match.group('basic_item')
#                     change = match.group('to_change')

#                     for key, value in first_sheet_values.items():
#                         if basic == key:
#                             row = value[0]
#                             column = value[1]
#                             pos = cell_value(first_data, row, column).coordinate
#                             first_data[pos] = change
#                             first_data[pos].value
#     print(first_sheet_values)
#     third_name = copy()
#     first_workbook.save(third_name)


def sheets_to_compare(wb):
    first_workbook_ask = input("Wollen sie ihr momentanes Spreadsheet als Ausgangsdatei verwenden? [Y/n]\n>  ").lower()
    clear()

    first_sheet = None
    second_sheet = None
    first_workbook = None
    second_workbook = None

    if first_workbook_ask == 'exit':
        exit()
    elif first_workbook_ask == 'n':

        first_workbook = ask_workbook("""Geben sie ihren Pfad zur Datei an oder ziehen sie ihn rein per drag and drop(es sollte eine Exel datei sein also z.B. .xlsx.  
Das Spreadsheet was sie angeben wird als Ausgangsdatei verwendet!\n>  """)
        clear()
        first_sheet = get_sheet(first_workbook)
        clear()
    else:
        first_workbook = wb
        first_sheet = get_sheet(first_workbook)
        clear()

    second_workbook_ask = input("Wollen sie eine andere Datei zum Vergleich nehmen? [Y/n]\n>  ")
    clear()

    if second_workbook_ask == 'exit':
        exit()
    elif second_workbook_ask == 'n':
        second_sheet = get_sheet(first_workbook)
        clear()
    else:

        second_workbook = ask_workbook("""Geben sie ihren Pfad zur Datei an oder ziehen sie ihn rein per drag and drop(es sollte eine Exel datei sein z.B. .xlsx).\n
Das Spreadsheet was sie angeben wird als Vergleichsdatei verwendet!\n>  """)
        clear()
        second_sheet = get_sheet(second_workbook)
        clear()

    return first_sheet, second_sheet, first_workbook_ask, first_workbook


def position(sheet):
    position_data = input("Geben sie die Positon ihrer Zelle an wie: 'Reihe, Spalte' (z.B. 2,1).\n>  ").lower()

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
            return cell_value(sheet, match.group('row'), match.group('column')).value
    else:
        exit()


def compare(wb):
    # Ermittelt die gewünschten Sheets und workbooks
    data = sheets_to_compare(wb)
    first_data = data[0]
    second_data = data[1]
    first_workbook_name = data[2]
    first_workbook = data[3]

    # Ermittelt Maximale Reihe und Spalte des ersten Sheets†
    first_max_sheet = max_sheet(first_data)
    first_max_row = first_max_sheet[0]
    first_max_column = first_max_sheet[1]

    # Ermittelt Maximale Reihe und Spalte des zweiten Sheets
    second_max_sheet = max_sheet(second_data)
    second_max_row = second_max_sheet[0]
    second_max_column = second_max_sheet[1]

    # Ermittelt Werte der sheets
    first_values = get_values(first_data, first_max_row, first_max_column)[0]
    second_values = get_values(second_data, second_max_row, second_max_column)[0]

    # Delete first value
    del first_values[0]
    del second_values[0]

    # Compares the sheets
    compare_sheets(first_values, second_values, first_data, first_workbook)


def grid(values, longest):
    # Vorher raw_grid
    string_rows = []
    for rows_in_list in values:
        raw_temp = []
        for item in rows_in_list:
            value_length = longest + 1 - len(str(item))
            raw_temp.append(str(item) + " " * value_length + "|")
        string_rows.append(raw_temp)

    # Vorher grid
    for row in string_rows:
        string_temp = ""

        # Fügt die Reihen Zahlen hinzu
        string_temp += str(string_rows.index(row)) + "| "

        for i in range(0, len(row)):
            string_temp += row[i]
        print(string_temp)


def all(wb):
    # Fragt neu gewünschtes Sheet ab
    sheet = get_sheet(wb)
    clear()

    # Ermittelt Maximale Reihe und Spalte
    max_row_column = max_sheet(sheet)
    max_row = max_row_column[0]
    max_column = max_row_column[1]

    # Zwei Dimensionale Liste mit den den raw Row Werten
    values = get_values(sheet, max_row, max_column)[0]

    # Ermittelt die längste Zelle des Sheets
    longest = get_values(sheet, max_row, max_column)[1]

    # Verarbeitet die Werte zu einem Grid von Strings
    grid(values, longest)


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


def help():
    longest_key = 0
    funktionen = {
    "all" : "Gibt eine tabellarische Übersicht des ausgewählten Sheets wieder",
    "view": "Gibt den Inhalt einer gewünschten Zelle wieder",
    "compare": "Schreibt eine neue Datei um Datensätze von B auf A zu überschreiben falls es überschneidungen gibt.",
    "exit": "Velässt das Programm",
    "github": "https://github.com/Backerich/Spreadsheet"}

    for key, value in funktionen.items():
        if len(key) > longest_key:
            longest_key = len(key)

    for key, value in funktionen.items():
        print(key + " " * (longest_key - len(key)) + "| " + value)

    print("\n")
    ask_function = input("Welche Funktion möchten sie erklärt haben?\n>  ").lower()

    if ask_function == 'exit':
        exit()
    else:
        clear()
        print("Wird noch implimentiert.")


def menu(wb):
    clear()

    while True:
        # Alle möglichen Funktionen des Programmes als Menü
        print("Gib 'all' ein um einen überblick von deinem Spreadsheet zu erhalten.")
        print("Gib 'view' ein um bestimmt Zeilen zu inspizieren.")
        print("Gib 'compare' ein um Datensätze der ersten Datei mit Datensätzen einer zweiten zu ersetzen.")
        print("Gib 'edit' ein um dein workbook zu ändern.")
        print("Gib 'help' ein wenn du hilfe brauchst.")
        print("Gib 'exit' ein um das Programm zu verlassen.")

        # Der Userinput um auf die einzelnen Menüpunkte zuzugreifen:
        # user_input = input(">  ").lower()

        # DEBUG: Als default der Eingabe:
        user_input = 'compare'
        
        if user_input == 'exit':
            exit()
        elif user_input == 'view':
            # Falls Input 'view' Abfrage vom view um einzelnen Zelle Inhalte zu betrachen
            clear()

            #
            print(view(wb))

            # Abrage ob das Programm beendet werden soll
            continue_request(wb, "\nWollen sie weiter machen [Y/n]? \n>  ")
        elif user_input == 'edit':
            # Falls input 'edit' wird abgefragt ob der User das Workbook ändern will bzw. in welches
            clear()

            # 
            print("Ändere dein workbook.\n")

            # Ersetzt das aktuelle Workbook mit dem neuen
            wb = ask_workbook("Geben sie ihren Pfad zur Datei an oder ziehen sie ihn rein per drag and drop(es sollte eine Exel datei sein also z.B. .xlsx.\n>  ")
            clear()
        elif user_input == 'all':
            # Falls input 'all' wird eine tabellarische Übersicht des ausgewählten Sheets angezeigt
            clear()

            # Zeigt alle Inhalte des Sheets an(eine Übersicht)
            all(wb)

            # Abrage ob das Programm beendet werden soll
            continue_request(wb , "\nWollen sie weiter machen [Y/n]? \n>  ")
        elif user_input == 'compare':
            # Falls input 'compare' wird Datei A mit Datei B verglichen und einstimmige Komponenten ersetzt
            clear()

            #
            compare(wb)

            # Abrage ob das Programm beendet werden soll
            continue_request(wb , "\nWollen sie weiter machen [Y/n]? \n>  ")
        elif user_input == 'help':
            clear()

            # Erläuterung der Eingabe optionen
            help()

            # Abrage ob das Programm beendet werden soll
            continue_request(wb, "\nWollen sie weiter machen [Y/n]? \n>  ")
        else:
            # Fehlermeldung falls es diesen Menüpunkt nicht gibt und läuft die Schleife nochmal durch
            clear()

            #
            print("Den Befehl '{}'' gibt es nicht. Versuchen sie es noch einmal.\n".format(user_input))

            #
            continue


def main():
    clear()

    # DEBUG: Default Workbbok zum Debugen
    wb = openpyxl.load_workbook('Example/example.xlsx')
    # Fragt Workbook ab
    # wb = ask_workbook("Geben sie ihren Pfad zur Datei an oder ziehen sie ihn rein per drag and drop(es sollte eine Exel datei sein also z.B. .xlsx.\n>  ")

    menu(wb)


if __name__ == "__main__":
    main()

