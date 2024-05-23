import openpyxl
import tkinter
from tkinter import filedialog
from openpyxl.reader.excel import load_workbook

root = tkinter.Tk()
root.withdraw()

#wb = load_workbook('BTT_Template.xlsx')
wb = load_workbook(filedialog.askopenfilename(filetypes=[("Excel files", '*.xlsx *.xls')]))

file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])

all_tab_data = {}


def copy_range(start_col, start_row, end_col, end_row, sheet):
    range_sel = []
    for i in range(start_row, end_row + 1, 1):
        row_sel = []
        for j in range(start_col, end_col + 1, 1):
            row_sel.append(sheet.cell(row=i, column=j).value)
        range_sel.append(row_sel)
    return range_sel


def paste_range(start_col, start_row, end_col, end_row, sheet_receiving, copied_data):
    count_row = 0
    for i in range(start_row, end_row + 1, 1):
        count_col = 0
        for j in range(start_col, end_col + 1, 1):
            try:
                sheet_receiving.cell(row=i, column=j).value = copied_data[count_row][count_col]
            except Exception as e:
                #sheet_receiving.cell(row=i, column=j).value = None
                continue
            count_col += 1
        count_row += 1


for path in file_paths:
    wb_tmp = load_workbook(path)
    try:
        wb_tmp.remove(wb_tmp["Übersicht"])
        wb_tmp.remove(wb_tmp["Quercheck Transaktionen"])
    except Exception as e:
        print(str(e))
    tab_data = {}

    for sheet_name in wb_tmp.sheetnames:    # TODO: was ist mit sheets, die voneinander getrennte Tabellen haben?
        tab_data[sheet_name] = {
            "start_col": wb_tmp[sheet_name].min_column,
            "start_row": wb_tmp[sheet_name].min_row,
            "end_col": wb_tmp[sheet_name].max_column,
            "end_row": wb_tmp[sheet_name].max_row
        }
    all_tab_data[path] = tab_data

for file, data in all_tab_data.items():
    wb_tmp = load_workbook(file)
    for sheet_name, dim in data.items():
        if sheet_name == 'BTT':
            sel_range = copy_range(dim["start_col"], dim["start_row"]+2, dim["end_col"], dim["end_row"], wb_tmp[sheet_name])
        else:
            sel_range = copy_range(dim["start_col"], dim["start_row"]+1, dim["end_col"], dim["end_row"],
                                   wb_tmp[sheet_name])
        paste_range(wb[sheet_name].min_column, wb[sheet_name].max_row+1, wb[sheet_name].max_column, wb[sheet_name].max_row+dim["end_row"], wb[sheet_name], sel_range) # TODO: die dinger stimmen nicht


# TODO:
#   - paste_range anpassen
#   - Duplikate behandeln
#   - Formatierung erweitern
#   - Drop down ertweitern  -> DataValidation <https://openpyxl.readthedocs.io/en/2.5/validation.html>, <https://stackoverflow.com/questions/51497731/openpyxl-is-it-possible-to-create-a-dropdown-menu-in-an-excel-sheet>
#   - Formeln überprüfen -> Sheet neu laden?


wb.save('output.xlsx')
