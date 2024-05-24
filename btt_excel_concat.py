import openpyxl
import tkinter
from tkinter import filedialog
from openpyxl.reader.excel import load_workbook

root = tkinter.Tk()
root.withdraw()

wb = load_workbook('BTT_Template.xlsx')
#wb = load_workbook(filedialog.askopenfilename(filetypes=[("Excel files", '*.xlsx *.xls')]))

file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])

all_tab_data = {}


def get_max_row(sheet, start_col, end_col):
    max_row = 0
    if start_col == end_col:
        for row in range(1, sheet.max_row+1):
            if sheet.cell(row=row, column=start_col).value is not None:
                max_row = max(max_row, row)
    else:
        for col in range(start_col, end_col, 1):
            for row in range(1, sheet.max_row+1):
                if sheet.cell(row=row, column=col).value is not None:
                    max_row = max(max_row, row)
    return max_row


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
                #sheet_receiving.cell(row=i, column=j).value = None TODO: braucht man das?
                continue
            count_col += 1
        count_row += 1


for path in file_paths:
    wb_tmp = load_workbook(path)
    try:
        wb_tmp.remove(wb_tmp["Quercheck Transaktionen"])
    except Exception as e:
        print(str(e))

    tab_data = {}

    for sheet_name in wb_tmp.sheetnames:
        if sheet_name == "Übersicht":
            tab_data[sheet_name] = [
                {"start_col": 1, "start_row": 4, "end_col": 2, "end_row": get_max_row(wb_tmp[sheet_name], 1, 2)},
                # [A:B]
                {"start_col": 5, "start_row": 2, "end_col": 8, "end_row": get_max_row(wb_tmp[sheet_name], 5, 8)}
                # [E:H]
            ]
        elif sheet_name == "BTT":
            tab_data[sheet_name] = [
                {"start_col": 1, "start_row": 3, "end_col": wb_tmp[sheet_name].max_column, "end_row": wb_tmp[sheet_name].max_row}
            ]
        elif sheet_name == "BPML":
            tab_data[sheet_name] = [
                {"start_col": 1, "start_row": 2, "end_col": 4, "end_row": get_max_row(wb_tmp[sheet_name], 1, 4)}, # [A:D]
                {"start_col": 6, "start_row": 2, "end_col": 10, "end_row": get_max_row(wb_tmp[sheet_name], 6, 10)}  # [F:J]
            ]
        elif sheet_name == "Datengrundlage adesso":
            tab_data[sheet_name] = [
                {"start_col": 1, "start_row": 2, "end_col": 3, "end_row": get_max_row(wb_tmp[sheet_name], 1, 3)},
                # [A:C]
                {"start_col": 5, "start_row": 2, "end_col": 5, "end_row": get_max_row(wb_tmp[sheet_name], 5, 5)},
                # [E]
                {"start_col": 7, "start_row": 2, "end_col": 7, "end_row": get_max_row(wb_tmp[sheet_name], 7, 7)},
                # [G]
                {"start_col": 9, "start_row": 2, "end_col": 9, "end_row": get_max_row(wb_tmp[sheet_name], 9, 9)},
                # [I]
                {"start_col": 11, "start_row": 2, "end_col": 11, "end_row": get_max_row(wb_tmp[sheet_name], 11, 11)}
                # [I]
            ]
        else:
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
        if isinstance(dim, list):
            for d in dim:
                sel_range = copy_range(d["start_col"], d["start_row"], d["end_col"], d["end_row"], wb_tmp[sheet_name])
                paste_range(d["start_col"], get_max_row(wb[sheet_name], d["start_col"], d["end_col"])+1, d["end_col"], wb[sheet_name].max_row+d["end_row"], wb[sheet_name], sel_range)
        else:
            sel_range = copy_range(dim["start_col"], dim["start_row"]+1, dim["end_col"], dim["end_row"], wb_tmp[sheet_name])
            paste_range(wb[sheet_name].min_column, get_max_row(wb[sheet_name], dim["start_col"], dim["end_col"])+1, wb[sheet_name].max_column,
                        wb[sheet_name].max_row + dim["end_row"], wb[sheet_name], sel_range)


# TODO:
#   - Duplikate behandeln?
#   - Formatierung erweitern
#   - Drop down ertweitern  -> DataValidation <https://openpyxl.readthedocs.io/en/2.5/validation.html>, <https://stackoverflow.com/questions/51497731/openpyxl-is-it-possible-to-create-a-dropdown-menu-in-an-excel-sheet>
#   - Formeln überprüfen -> Sheet neu laden?
#   - schön machen


wb.save('output.xlsx')
