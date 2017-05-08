#!/usr/bin/env python3

import openpyxl

import sys
import os
import re
# import time
# import traceback


###############################################################################
# PROBLEME: -view unerwarteter error -> siehe unten (nicht wieder aufgetreten)
#           -
# IMPLIMENTIEREN: -make the cotinue func universal
#                 -aktive Sheet
#                 -übersicht
#                 -
###############################################################################


def exit():
    return sys.exit()


def clear():
    os.system("cls" if os.name == "nt" else "clear")


def continue_loop(wb):
    continue_loop = input("\nWollen sie weiter machen [Y/n]? \n>  ")
    if (continue_loop == 'n') or (continue_loop == 'exit'):
        exit()
    else:
        menu(wb)


def what_sheet(wb):
    this_sheet = input("Welches Sheet willst du convertieren? \n>  ") 

    if this_sheet == 'exit':
        exit()
    else:
        try:
            return wb.get_sheet_by_name(this_sheet)
        except KeyError:
                clear()
                print("\nDas Sheet {} existiert nicht. Versuchen sie es nochmal.\n".format(this_sheet))
                view(wb)


def list_sheets(wb):
    sheets = wb.get_sheet_names()
    counter = 1

    for string in sheets:
        print("Sheet Nummer {}: {}".format(counter, string))
        counter += 1


def get_sheet(wb):
    list_sheets(wb)
    return what_sheet(wb)


def cell_value(sheet, row_first, column_first):
    return sheet.cell(row=int(row_first), column=int(column_first)).value


def all(wb):
    sheet = get_sheet(wb)
    clear()

    row_count = sheet.max_row
    column_count = sheet.max_column
    all_data = {}

    # header immer selbe row und ander column
    for column in range(1 , column_count + 1):
        for row in range(1, row_count + 1):
            values = cell_value(sheet, row, column)
            if values != None:
                print(values)
                # all_data = {row , {}}
                print(all_data)
            else:
                print("\n")


def position(sheet):
    position_data = input("Geben sie die Positon ihrer Zelle an wie: 'Reihe, Spalte' (z.B. 2,1).\n>  ")

    if position_data != 'exit':
        clear()

        line = re.compile(r'''
        ^(?P<row>[\d]+),\s*
        (?P<column>[\d]+)$
        ''', re.X|re.M)

        for match in line.finditer(position_data):
            return cell_value(sheet, match.group('row'), match.group('column'))
    else:
        exit()


def view(wb):
    sheet = get_sheet(wb)
    clear()

    position_data = position(sheet)

    if position_data != None:
        return "Ihre Zelle beinhaltet '{}'.".format(position_data)
    else:
        return "Diese Zelle gibt es nicht bitte versuchen sie es nochmal."


def ask_workbook():
    workbook_input = input("Gib deinen Pfad zur Datei an oder zieh ihn rein per drag and drop.\n>  ")
    try:
        return openpyxl.load_workbook(workbook_input)
    except FileNotFoundError:
        clear()
        print("Diser Pfad exsistiert nicht. Versuch es nochmal.")
        ask_workbook()

def menu(wb):
    clear()
    while True:
        print("Gib 'exit' ein um das Programm zu verlassen.")
        print("Gib 'view' ein um bestimmt Zeilen zu inspizieren.")
        print("Gib 'edit' ein um dein workbook zu ändern.")
        print("Gib 'all' ein um einen überblick von deinem Spread sheet zu erhalten.")
        user_input = input(">  ").lower()
        
        if user_input == 'exit':
            exit()
        elif user_input == 'view':
            clear()
            print(view(wb))
            continue_loop(wb)
        elif user_input == 'edit':
            clear()
            print("Ändere dein workbook.\n")
            wb = ask_workbook()
            clear()
        elif user_input == 'all':
            clear()
            all(wb) # muss noch implimentiert werden
            continue_loop(wb)
        else:
            clear()
            print("Den Befehl '{}'' gibt es nicht. Versuchen sie es noch einmal.\n".format(user_input))
            continue
def main():
    clear()
    wb = openpyxl.load_workbook('example_Kopie.xlsx')
    # wb = ask_workbook()
    menu(wb)

if __name__ == "__main__":
    main()




