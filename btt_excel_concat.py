import datetime
import os.path

import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
import pandas as pd
import tkinter
from tkinter import filedialog

root = tkinter.Tk()
root.withdraw()
sel_cols = ['Lfd Nr.\n(automatisch)',
            'Hauptprozess\n(Pflichtauswahl)',
            'Subprozess\n(optionale Auswahl)',
            'Prozessschritt / Funktionsname (Freitext - Pflicht)',
            'Verantwortliches TP\n(automatisch)',
            'Manuelle Änderung des Verantwortliches TP\n(Auswahl - bei Bedarf)',
            'Info zu OEen\n(Freitext - bei Bedarf)',
            'SAP-Modul\n(Pflichtauswahl)',
            'Verwendete Transaktion (Pflichtauswahl)',
            'Transaktions-name (automatisch)',
            'Zugehörige Transaktionen (Freitext - optional)',
            'Verwendete \nFiori App (Freitext - optional)',
            'Z-Entwicklung zur Transaktion\n(Freitext - optional)',
            'Verwendetes Addon\n(Freitext - optional)',
            'Digital signiert\n(Pflichtauswahl)',
            'Verwendeter Workflow\n(Freitext falls relevant)',
            'Verwendete Business Function\n(Freitext falls relevant)',
            'Verwendete Schnittstelle\n(optionale Auswahl)',
            'Weitere Schnittstellen (Freitext - optional)',
            'Art des Outputs\n(Pflichtauswahl)',
            'Verwendetes Formular\n(Auswahl falls relevant)',
            'technischer Formularname (automatisch)',
            'Verwendetes anderes\nOutputmedium \n(Freitext falls relevant)',
            'Org Management Relevanz\n(Pflichtauswahl)',
            'Anmerkungen\n(Freitext - optional)',
            'Priorität\n(Pflichtauswahl)']
file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
concat_df = []
empty_cols = []
existing_df = pd.DataFrame()
da = pd.DataFrame()
existing_row_idx = []

if os.path.isfile('output.xlsx'):
    existing_df = pd.read_excel('output.xlsx', sheet_name='BTT', header=0, index_col=None)
    existing_row_idx = list(existing_df.index)

for path in file_paths:
    try:
        da = pd.read_excel(open(path, 'rb'), sheet_name="BTT", header=0, index_col=None, usecols=sel_cols, skiprows=1)
        print("Datei", path, "erfolgreich gelesen.")

        # da.columns = da.columns.str.split('\n').str[0]
        # da.columns = [re.sub(r'\([^)]*\)', '', col) for col in da.columns]
        # print(da.columns)

        da['Quelldatei'] = os.path.basename(path)
        da['Datum'] = datetime.date.today().strftime("%d.%m.%Y")

    except Exception as e:
        print("Fehler beim Lesen der Datei", path, ":", str(e))
        try:
            da = pd.read_excel(open(path, 'rb'), sheet_name="BTT", header=0, index_col=None, skiprows=1)
            for col in sel_cols:
                if col not in da.columns:
                    da[col] = None
            da = da.drop(columns=[col for col in da.columns if col not in sel_cols])

            da['Quelldatei'] = os.path.basename(path)
            da['Datum'] = datetime.date.today().strftime("%d.%m.%y")

        except Exception as e:
            print(str(e))

    concat_df.append(da)

df = pd.concat(concat_df)
df = pd.concat([existing_df, df]).drop_duplicates()

if not existing_df.empty:
    new_da = df[~df.isin(existing_df.to_dict('list')).all(1)]

with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, index=False, header=True, sheet_name='BTT')

    wb = writer.book
    ws = wb['BTT']

    for col in df.columns:
        if df[col].dropna().empty:
            empty_cols.append(col)

    for col in empty_cols:
        idx = df.columns.get_loc(col)
        ws.column_dimensions[ws.cell(row=1, column=idx + 1).column_letter].hidden = True

    max_width = 30
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except Exception as e:
                pass
        adjusted_width = min((max_length + 2) * 1.2, max_width)
        ws.column_dimensions[column].width = adjusted_width

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        if row_idx == 1:
            for cell in row:
                cell.alignment = Alignment(wrapText=True)
            continue
        for cell in row:
            cell.fill = openpyxl.styles.PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            cell.font = Font(color="007B74")
            cell.border = Border(left=Side(border_style="thin", color="7F7F7F"),
                                 right=Side(border_style="thin", color="7F7F7F"),
                                 top=Side(border_style="thin", color="7F7F7F"),
                                 bottom=Side(border_style="thin", color="7F7F7F"))

    for idx in existing_row_idx:
        row_num = idx + 2
        for cell in ws[row_num]:
            cell.fill = openpyxl.styles.PatternFill(fill_type=None)
            cell.font = Font(color="000000")
            cell.border = None
