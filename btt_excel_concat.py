import datetime
import os.path
from copy import copy

import openpyxl
import tkinter
from tkinter import filedialog

import pandas as pd
from openpyxl.reader.excel import load_workbook

'''
######################
        TODOS
######################
    - BTT Template muss angepasst werden
        - Sverweise, Formeln als solche übernehmen, NICHT die Werte in den Zellen
        - erst die Sheets mit den Wertehilfen konsolidieren
        - dann die BTT mit Werten füllen, auch hier die Formeln/Sverweise mit übernehmen 

'''

root = tkinter.Tk()
root.withdraw()
file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])

#wb = load_workbook('BTT_Template.xlsx')
wb = load_workbook(filedialog.askopenfilename(filetypes=[("Excel files", '*.xlsx *.xls')]))


def equal_ws(ws1, ws2, sheet_name):
    idx = []
    if ws1.max_row != ws2.max_row or ws1.max_column != ws2.max_column:
        return False

    if sheet_name == 'Übersicht':
        idx = [0, 4, 5]  # es fehlen die Spalten B, G, H
    elif sheet_name == 'BPML':
        idx = [0, 1, 2, 5, 6, 7]  # es fehlen die Spalten D, I, J
    elif sheet_name == 'Transaktionen':
        idx = [0, 1, 2, 3, 4, 6]  # es fehlen die Spalten F
    elif sheet_name == 'Formulare':
        idx = [0, 1]  # es fehlt Spalte C
    elif sheet_name == 'Schnittstellen':
        idx = [7, 8]  # es fehlt Spalte J -> was ist mit A:G????
    elif sheet_name == 'Datengrundlage adesso':  # TODO: immer gleich?
        idx = [0, 1, 4, 6, 8, 10]  # es fehlt die Spalte C

    for row1, row2 in zip(ws1.iter_rows(values_only=True), ws2.iter_rows(values_only=True)):
        for i in idx:
            cell1, cell2 = row1[i], row2[i]
            if cell1 != cell2:
                return False
    return True


for path in file_paths:
    wb_tmp = load_workbook(path)
    try:
        for sheet in wb.sheetnames:
            try:
                ws_tmp = wb_tmp[sheet]

                if sheet == 'BTT':
                    # TODO sind A2:_ ; B2:_ ; C2:_ identisch?
                    if not equal_ws(ws_tmp, wb[sheet], sheet):  # TODO: ganzes sheet wird noch überprüft
                        for row in ws_tmp:
                            wb[sheet].append(row)
                        #wb[sheet].append(ws_tmp)
                        print("sheets", sheet, "not equal")
                elif sheet == 'BPML':
                    # TODO sind A2:_ ; B2:_ ; C2:_ identisch?
                    if not equal_ws(ws_tmp, wb[sheet], sheet):  # TODO: ganzes sheet wird noch überprüft
                        for row in ws_tmp:
                            wb[sheet].append(row)
                        #wb[sheet].append(ws_tmp)
                        print("sheets", sheet, "not equal")
                elif sheet == 'Transaktionen':
                    # TODO sind A2:_ ; B2:_ ; C2:_ ; E2:_ identisch?
                    # sind D2:_ identisch? wenn nein, dann das größere nehmen?
                    if not equal_ws(ws_tmp, wb[sheet], sheet):
                        for row in ws_tmp:
                            wb[sheet].append(row)
                        #wb[sheet].append(ws_tmp)
                        print("sheets", sheet, "not equal")
                elif sheet == 'Formulare':
                    # TODO sind A2:_ ; B2:_ identisch?
                    if not equal_ws(ws_tmp, wb[sheet], sheet):
                        for row in ws_tmp:
                            wb[sheet].append(row)
                        #wb[sheet].append(ws_tmp)
                        print("sheets", sheet, "not equal")
                elif sheet == 'Schnittstellen':
                    # TODO sind H,I,J2 gleich; was ist mit A:F??
                    if not equal_ws(ws_tmp, wb[sheet], sheet):
                        for row in ws_tmp:
                            wb[sheet].append(row)
                        #wb[sheet].append(ws_tmp)
                        print("sheets", sheet, "not equal")

            except Exception as e:
                print("Fehler beim Lesen der Mappe", sheet, "in Datei", path, ":", str(e))
    except Exception as e:
        print("Fehler beim Lesen der Datei", path, ":", str(e))

wb.save('output.xlsx')
