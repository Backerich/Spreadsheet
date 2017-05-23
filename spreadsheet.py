#!/usr/bin/env python3

import openpyxl

import sys
import os
import re

strings_german = []
strings_englisch = []

def exit():
    # Verlässt das Programm
    return sys.exit()


def clear():
    # Räumt den Terminal Output auf
    os.system("cls" if os.name == "nt" else "clear")


def ask_workbook(input_string):
    string = strings_german[6] + input_string + "\n" + "> "
    workbook_input = input(string).lower()
    if workbook_input == strings_german[0]:
        exit()

    else:
        try:
            # Falls es den Path gibt verwendet er nun dieses Workbook
            return openpyxl.load_workbook(workbook_input)
        except FileNotFoundError:
            # Falls nicht wird eine Fehlermeldung ausgegeben und nochmal abgefragt
            clear()
            print(strings_german[22] + "\n")
            ask_workbook(input_string)


def continue_request(wb):
    # Abfrage ob Programm beendet werden soll oder weiter laufen soll
    continue_loop = input("\n" + strings_german[7] + "\n" + "> ").lower()
    if (continue_loop == "n") or (continue_loop == strings_german[0]):
        exit()

    else:
        # Wenn nicht 'exit' geht es wieder zum Ausgangsmenü zurück
        menu(wb)


def what_sheet(wb):
    this_sheet = input(strings_german[21] + "\n" + "> ").lower()

    if this_sheet == strings_german[0]:
        exit()

    elif this_sheet == "":
        # Wenn nur Enter gedrückt wird wird das aktive Sheet als default genommen
        return wb.get_active_sheet() 

    else: 
        try:
            # Falls der Inputname existiert wird das ausgewählte Sheet benutzt
            return wb.get_sheet_by_name(this_sheet)
        except KeyError:
                # Wenn nicht wird das Terminal aufgeräumt, eine Fehlermeldung ausgegeben und es wird erneut view abgefragt
                clear()
                print("\n" + strings_german[38] + "\n".format(this_sheet))
                get_sheet(wb) 


def list_sheets(wb):
    # Listet alle Sheets in dem workbook auf
    try:
        # Absicherung durch unerwarteten Bug:
        # ERROR: AttributeError: 'NoneType' object has no attribute 'get_sheet_names'
        sheets = wb.get_sheet_names()
    except AttributeError:
        clear()
        print(strings_german[20])
        menu(wb)
    counter = 1

    for string in sheets:
        print(strings_german[37].format(counter, string))

        # counter zur Orientierung des Endusers
        counter += 1


def max_sheet(sheet):
    try:
        # Falls es eine Reihe und Spalte gibt wird die höchste wiedergegeben
        return sheet.max_row, sheet.max_column
    except AttributeError:
        # Falls es keine Reihe oder Spalte gibt wird eine Fahlermeldung ausgegeben
        print(strings_german[19])


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

    # Fügt die Spalten beschreibung hinzu
    all_raw_rows.append(i for i in  range(1, row_count + 1))

    for row in range(1, row_count + 1):
        innerlist = []

        for column in range(1 , column_count + 1):
            # Ermittelt Zellen Inhalt
            values = cell_value(sheet, row, column).value

            if values != None:
                innerlist.append(values)

                # Ermittelt längsten Wert
                if len(values) > longest:
                    longest = len(values)

            else:
                # Abgrenzungen
                innerlist.append("")

        # Fügt alle zusammen
        all_raw_rows.append(innerlist)
    return all_raw_rows, longest


def copy():
    # Fragt nach dem Namen der Datei C
    print(strings_german[18])
    copy_name = input(strings_german[17] + "\n" + "> ").lower().strip()

    if copy_name == strings_german[0]:
        exit()

    else:

        # Gibt den Pfad zurück
        return "Example/" + copy_name


def compare_sheets(first_values, second_values, first_data, first_workbook):
    first_sheet_values = []

    for first_row in first_values:
        for first_item in first_row:
            if first_item != "":

                # Teilt die Ausgangszelle per Leerzeichen
                item_list = first_item.split(" ") # RE eher angebracht

                for item_raw in item_list:
                    # Filtert Kommata heraus
                    item = item_raw.replace(",", "") # RE eher angebracht

                    # Ermittelt die Postion der Reihe, der Spalte und die Position im String
                    position_row = first_values.index(first_row)
                    position_column = first_row.index(first_item)
                    position_string = item_list.index(item_raw)

                    # Gibt das Objekt mit den Koordinaten weiter
                    first_sheet_values.append([item, (position_row + 1, position_column + 1, position_string)])

                # Setzt "done" ans ende des benutzten Objektes damit selbes Item mit anderen Koordinaten existieren kann
                first_row[position_column] += "done"

    for second_row in second_values:
        for second_item in second_row:

            # Wenn item Leer ist weiter machen
            if second_item == "":
                pass

            else:
                # Löscht Leerzeichen und teilt Werte auf durch "="
                second_item = second_item.replace(" ", "").split("=")

                for cell in first_sheet_values: # iter nicht immer über alle
                    item = cell[0]

                    # Wenn Übereinstimmung des Wertes und des Vergleichsobjekt
                    if second_item[0] == item:

                        # Ermittelt die position
                        row = cell[1][0]
                        column = cell[1][1]
                        pos = cell_value(first_data, row, column).coordinate

                        # Ermittelt Inhalt der Zelle
                        value = first_data[pos].value

                        ####################################################################
                        # Gucken ob 'find' nicht den oberen Abschnitt der Funktion ersetzt #
                        ####################################################################

                        # Sucht Vergleichs datei in Ausgangsdatei
                        index = value.find(second_item[0])

                        try: # Checken ob Errorhandling nötig
                            # Löscht Vergleichsdatei in Ausgangsdatie
                            value = value.replace(second_item[0], "")

                            # Setzt String wieder zusammen
                            value = value[:index] + second_item[1] + value[index:]

                            first_data[pos] = value
                        except IndexError:
                            pass
    # Erstellt kopie von von Datei A und speichert es in Datei C
    third_name = copy()
    first_workbook.save(third_name)


def sheets_to_compare(wb): # Überlegung ob nötig
    user_workbook = input(strings_german[16] + "\n" + "> ").lower()

    sheet = None
    workbook = None

    if user_workbook == strings_german[0]:
        exit()

    elif user_workbook == "n":
        # Erfolgt wenn anderes Workbook genutzt werden soll und ermittelt das Worksheet
        workbook = ask_workbook("\n" + strings_german[15])
        clear()
        sheet = get_sheet(workbook)
        clear()

    else:
        # Erfolgt wenn selbes Workbook genutzt werden soll und ermittelt das Worksheet
        workbook = wb
        clear()
        sheet = get_sheet(workbook)
        clear()

    return sheet, user_workbook, workbook


def position(sheet):
    position_data = input(strings_german[12] + "\n" + "> ").lower()

    if position_data == strings_german[0]:
        exit()

    else:
        clear()

        # RE: Muster zum filtern der Koordinaten 
        line = re.compile(r'''
        ^(?P<row>[\d]+),\s*
        (?P<column>[\d]+)$
        ''', re.X|re.M)

        # Ermittelt die Werte durch RE und setzt sie in die Funktion ein um den Zelleninhalt heraus zu bekommen
        for match in line.finditer(position_data):
            return cell_value(sheet, match.group("row"), match.group("column")).value


def compare(wb):
    # Ermittelt die gewünschten Worksheets und Workbooks
    # Erstes Worksheet
    data_first = sheets_to_compare(wb)
    first_data = data_first[0]
    first_workbook_name = data_first[1]
    first_workbook = data_first[2]

    # Zweites Worksheet
    data_second = sheets_to_compare(wb) # BUG: Falsche Strings
    second_data = data_second[0]

    # Ermittelt Maximale Reihe und Spalte des ersten Worksheets
    first_max_sheet = max_sheet(first_data)
    first_max_row = first_max_sheet[0]
    first_max_column = first_max_sheet[1]

    # Ermittelt Maximale Reihe und Spalte des zweiten Worksheets
    second_max_sheet = max_sheet(second_data)
    second_max_row = second_max_sheet[0]
    second_max_column = second_max_sheet[1]

    # Ermittelt die Werte der Worksheets
    first_values = get_values(first_data, first_max_row, first_max_column)[0]
    second_values = get_values(second_data, second_max_row, second_max_column)[0]

    # Löschen des ersten Wertes
    del first_values[0]
    del second_values[0]

    # Vergleicht die Worksheets und gibt die Ergebnisse aus
    compare_sheets(first_values, second_values, first_data, first_workbook)


def grid(values, longest):
    # Convertiert die Werte in symetrische Reihen aus Strings
    string_rows = []
    for rows_in_list in values:
        raw_temp = []

        for item in rows_in_list:
            # Ermittelt die länge der spezifischen Zelle
            value_length = longest + 1 - len(str(item))

            # Setzt die Reihe zusammen
            raw_temp.append(str(item) + " " * value_length + "|")

        # Fügt die Reihe dem Raster hinzu
        string_rows.append(raw_temp)

    # Fügt Reihenzahlen hinzu und gibt die Tabelle aus
    for row in string_rows:
        string_temp = ""

        # Ermittelt die Reihenzahl
        string_temp += str(string_rows.index(row)) + "| "

        for i in range(0, len(row)):
            # Fügt die Reihenzahlen und die restliche Reihe zusammen
            string_temp += row[i]

        # Gibt die Tabelle aus
        print(string_temp)


def all(wb):
    # Fragt das Worksheet ab
    sheet = get_sheet(wb)
    clear()

    # Ermittelt Maximale Reihe und Spalte
    max_row_column = max_sheet(sheet)
    max_row = max_row_column[0]
    max_column = max_row_column[1]

    # Ermittelt die Rohwerte der Reihen
    values = get_values(sheet, max_row, max_column)[0]

    # Ermittelt die längste Zelle des Worksheets
    longest = get_values(sheet, max_row, max_column)[1]

    # Verarbeitet die Werte zu einem Grid von Strings
    grid(values, longest)


def view(wb):
    # Fragt das Worksheet ab
    sheet = get_sheet(wb)
    clear()

    # Ermittelt den Wert der Zelle
    position_data = position(sheet)

    if position_data != None:
        # Gibt den Inhalt der Zelle aus
        return strings_german[36].format(position_data)
    else:
        # ERROR: Wenn der Inhalt None ist
        return strings_german[11]


def help():
    # Die einzelnen Funktionen und ihre Beschreibung
    funktionen = {
    strings_german[5] : strings_german[31],
    strings_german[4]: strings_german[32],
    strings_german[3]: strings_german[33],
    strings_german[0]: strings_german[34],
    strings_german[30]: "https://github.com/Backerich/Spreadsheet"}

    # Ermittelt den längsten Key
    longest_key = 0
    for key, value in funktionen.items():
        if len(key) > longest_key:
            longest_key = len(key)

    # Gibt die Tabelle aus mit den Funktionsnamen und Beschreibungen
    for key, value in funktionen.items():
        print(key + " " * (longest_key - len(key)) + "| " + value)

    # Fragt nach der Funktion die weiter Erläutert werden soll
    ask_function = input("\n" + strings_german[10] + "\n" + "> ").lower()

    if ask_function == strings_german[0]:
        exit()

    else:
        # Gibt theoretisch die weitere Beschreibung aus
        clear()
        print(strings_german[9])


def menu(wb):
    # Das Ausgangsmenü um auf alle Funktionen zugreifen zu können
    clear()

    while True:
        # Alle möglichen Funktionen des Programmes
        print(strings_german[24].format(strings_german[5]))
        print(strings_german[25].format(strings_german[4]))
        print(strings_german[26].format(strings_german[3]))
        print(strings_german[27].format(strings_german[2]))
        print(strings_german[28].format(strings_german[1]))
        print(strings_german[29].format(strings_german[0]))

        # Benutzerinput für das Menü
        user_input = input("> ").lower()

        # DEBUG: Als default der Eingabe:
        # user_input = strings_german[3]
        
        if user_input == strings_german[0]:
            exit()

        elif user_input == strings_german[4]:
            # Zum betrachten des Wertes einer einzelner Zellen 
            clear()
            print(view(wb))
            continue_request(wb)

        elif user_input == strings_german[2]:
            # Zum ändern des aktuellen workbooks
            clear()
            print(strings_german[8] + "\n")
            wb = ask_workbook("")
            clear()

        elif user_input == strings_german[5]:
            # Gibt den gesammten Inhalt des Sheets in einer Tabelle aus
            clear()
            all(wb)
            continue_request(wb)

        elif user_input == strings_german[3]:
            # Datei A wird mit Datei B verglichen und einstimmige Komponenten in Datei C geschrieben und ersetzt
            clear()
            compare(wb)
            continue_request(wb)

        elif user_input == strings_german[1]:
            # Beschreibt die Funktionen des Programmes
            clear()
            help()
            continue_request(wb)

        else:
            # ERROR: Wenn es die Funktion nicht gibt
            clear()
            print(strings_german[23].format(user_input) + "\n")
            continue


def language():
    print("Type 'ger' for german.")
    print("Type 'eng' for englisch.")
    ask_language = input("Which language should your programm use?\n> ")

    if ask_language == 'exit':
        exit()

    elif ask_language == 'ger':
        with open("Languages/german", "r") as file:
            for line in file:
                full_line = line.replace("\n", "")
                strings_german.append(full_line)

    elif ask_language == 'eng':
        with open("Languages/englisch", "r") as file:
            for line in file:
                full_line = line.replace("\n", "")
                strings_englisch.append(full_line)

    else:
        clear()
        print("That language does not exist")
        language()


def main():
    clear()

    # Fragt Welche Sprache
    language()
    clear()

    # DEBUG: Default Workbbok zum Debuggen
    # wb = openpyxl.load_workbook('Example/example.xlsx')

    # Fragt nach dem workbook
    wb = ask_workbook("") # BUG: Wenn falsch kann compare nicht weiter arbeiten

    menu(wb)


if __name__ == "__main__":
    main()

