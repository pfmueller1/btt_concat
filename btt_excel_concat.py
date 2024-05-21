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
                sheet_receiving.cell(row=i, column=j).value = None
            count_col += 1
        count_row += 1


for path in file_paths:
    wb_tmp = load_workbook(path)
    tab_data = {}

    for sheet in wb_tmp:    # TODO: was ist mit sheets, die voneinander getrennte Tabellen haben?
        if sheet != 'Übersicht' and sheet != 'das andere halt':
            tab_data[sheet] = {
                "start_col": sheet.min_column,
                "start_row": sheet.min_row,
                "end_col": sheet.max_column,
                "end_row": sheet.max_row
            }
    all_tab_data[path] = tab_data

for file, data in all_tab_data.items():
    wb_tmp = load_workbook(file)
    for sheet, dim in data.items():
        sel_range = copy_range(dim["start_col"], dim["start_row"], dim["end_col"], dim["end_col"], sheet)
        paste_range(wb[sheet].min_column, dim["start_row"] + wb[sheet].max_row, wb[sheet].max_column, # TODO: wie Dimensionen des Templates behandeln?
                    dim["end_row"] + wb[sheet].max_row, wb[sheet], sel_range)

# ws.delete_rows(2)     TODO: index rutscht nach!
wb.save('output.xlsx')
