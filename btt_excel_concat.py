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


def equal_ws(ws1, ws2):
    if ws1.max_row != ws2.max_row or ws1.max_column != ws2.max_column:
        return False

    for row1, row2 in zip(ws1.iter_rows(values_only=True), ws2.iter_rows(values_only=True)):
        for cell1, cell2 in zip(row1, row2):
            if cell1 != cell2:
                return False
    return True


for path in file_paths:
    try:
        for sheet in wb.sheetnames:
            try:
                wb_tmp = load_workbook(path)
                ws_tmp = wb_tmp[sheet]
                '''                                       
                    - BTT: übernehme reine Wertefelder ohne weiteres
                        - Sonderfälle:
                            - Spalte A, E, J, V:
                                - Anpassung der Formel (Dateiname)?
                '''

                # Betrachtung der anderen sheets
                if sheet == 'Übersicht':
                    # TODO sind A4:_ und Formeln in B4:_ identisch?
                    if not equal_ws(ws_tmp, wb[sheet]): # TODO: ganzes sheet wird noch überprüft
                        pd.concat([pd.DataFrame(wb[sheet].values), pd.DataFrame(ws_tmp.values)],
                                  ignore_index=True).drop_duplicates()
                    continue
                elif sheet == 'BPML':
                    # TODO sind A2:_ ; B2:_ ; C2:_ identisch?
                    continue
                elif sheet == 'Transaktionen':
                    # TODO sind A2:_ ; B2:_ ; C2:_ ; E2:_ identisch?
                    # sind D2:_ identisch? wenn nein, dann das größere nehmen?
                    continue
                elif sheet == 'Formulare':
                    # TODO sind A2:_ ; B2:_ identisch?
                    continue
                elif sheet == 'Schnittstellen':
                    # TODO sind H,I,J2 gleich; was ist mit A:F??
                    continue
                elif sheet == 'Datengrundlage adesso':
                    # TODO vielleicht immer gleich??
                    continue

                if sheet != 'BTT' and ws_tmp.max_row != wb[sheet].max_row:
                    print(sheet)


            except Exception as e:
                print("Fehler beim Lesen der Mappe", sheet, "in Datei", path, ":", str(e))
    except Exception as e:
        print("Fehler beim Lesen der Datei", path, ":", str(e))

wb.save('output.xlsx')
