import openpyxl
import tkinter
from tkinter import filedialog
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

root = tkinter.Tk()
root.withdraw()

wb = load_workbook('BTT_Template.xlsx')
#wb = load_workbook(filedialog.askopenfilename(filetypes=[("Excel files", '*.xlsx *.xls')]))

file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])

all_tab_data = {}


def get_max_row(sheet, start_col, end_col):
    max_row = 0
    if start_col == end_col:
        for row in range(1, sheet.max_row + 1):
            if sheet.cell(row=row, column=start_col).value is not None:
                max_row = max(max_row, row)
    else:
        for col in range(start_col, end_col, 1):
            for row in range(1, sheet.max_row + 1):
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
                continue
            count_col += 1
        count_row += 1


def adj_table_dim(sheet, table_name):
    table = sheet.tables[table_name]
    ref = table.ref
    start_col, start_row = ref.split(':')[0][0], ref.split(':')[0][1:]
    end_col = ref.split(':')[1][0]
    end_row = get_max_row(sheet, column_index_from_string(start_col), column_index_from_string(end_col))
    new_ref = f"{start_col}{start_row}:{end_col}{end_row}"
    table.ref = new_ref


def add_dv(sheet):
    max_row = sheet.max_row

    for tab_name, cols in dv_lists.items():
        for col in cols:
            col_letter = col if col.isalpha() else get_column_letter(col)
            dv = DataValidation(
                type="list",
                formula1=f"{tab_name}",
                allow_blank=True,
                showDropDown=False,
                showInputMessage=True,
                showErrorMessage=True
            )
            range_str = f"${col_letter}$3:${col_letter}${max_row}"
            dv.add(range_str)
            sheet.add_data_validation(dv)


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
                {"start_col": 1, "start_row": 3, "end_col": wb_tmp[sheet_name].max_column,
                 "end_row": wb_tmp[sheet_name].max_row}
            ]
        elif sheet_name == "BPML":
            tab_data[sheet_name] = [
                {"start_col": 1, "start_row": 2, "end_col": 4, "end_row": get_max_row(wb_tmp[sheet_name], 1, 4)},
                # [A:D]
                {"start_col": 6, "start_row": 2, "end_col": 10, "end_row": get_max_row(wb_tmp[sheet_name], 6, 10)}
                # [F:J]
            ]
        elif sheet_name == "Datengrundlage adesso":
            tab_data[sheet_name] = [
                {"start_col": 1, "start_row": 2, "end_col": 3, "end_row": get_max_row(wb_tmp[sheet_name], 1, 3)},
                # [A:C]
                {"start_col": 5, "start_row": 2, "end_col": 5, "end_row": get_max_row(wb_tmp[sheet_name], 5, 5)},  # [E]
                {"start_col": 7, "start_row": 2, "end_col": 7, "end_row": get_max_row(wb_tmp[sheet_name], 7, 7)},  # [G]
                {"start_col": 9, "start_row": 2, "end_col": 9, "end_row": get_max_row(wb_tmp[sheet_name], 9, 9)},  # [I]
                {"start_col": 11, "start_row": 2, "end_col": 11, "end_row": get_max_row(wb_tmp[sheet_name], 11, 11)}
                # [K]
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
                paste_range(d["start_col"], get_max_row(wb[sheet_name], d["start_col"], d["end_col"]) + 1, d["end_col"],
                            wb[sheet_name].max_row + d["end_row"], wb[sheet_name], sel_range)
        else:
            sel_range = copy_range(dim["start_col"], dim["start_row"] + 1, dim["end_col"], dim["end_row"],
                                   wb_tmp[sheet_name])
            paste_range(wb[sheet_name].min_column, get_max_row(wb[sheet_name], dim["start_col"], dim["end_col"]) + 1,
                        wb[sheet_name].max_column,
                        wb[sheet_name].max_row + dim["end_row"], wb[sheet_name], sel_range)

        # expand table dimensions in template file
        if sheet_name == "Übersicht":
            adj_table_dim(wb[sheet_name], "Teilprojekte")
        elif sheet_name == "BPML":
            adj_table_dim(wb[sheet_name], "Hauptprozesse")
            adj_table_dim(wb[sheet_name], "BPML")
        elif sheet_name == "Transaktionen":
            adj_table_dim(wb[sheet_name], "Transaktionen")
        elif sheet_name == "Formulare":
            adj_table_dim(wb[sheet_name], "Formulare")
        elif sheet_name == "Schnittstellen":
            adj_table_dim(wb[sheet_name], "Schnittstelle_Klarname")
        elif sheet_name == "Datengrundlage adesso":
            adj_table_dim(wb[sheet_name], "Module")
            adj_table_dim(wb[sheet_name], "Prioritäten")
            adj_table_dim(wb[sheet_name], "Vorhanden?")
            adj_table_dim(wb[sheet_name], "Outputs")
            adj_table_dim(wb[sheet_name], "Interfaces")

# distribute drop down lists
dv_lists = {
    f'BPML!$A$2:$A${get_max_row(wb["BPML"], column_index_from_string("A"), column_index_from_string("A"))}': {'B'},
    f'=BPML!$F$2:$F${get_max_row(wb["BPML"], column_index_from_string("F"), column_index_from_string("F"))}': {'C'},
    f'=Transaktionen!$A$2:$A${get_max_row(wb["Transaktionen"], column_index_from_string("A"), column_index_from_string("A"))}': {
        'I'},
    f'=Schnittstellen!$H$2:$H${get_max_row(wb["Schnittstellen"], column_index_from_string("H"), column_index_from_string("H"))}': {
        'R'},
    f'=\'Datengrundlage adesso\'!$A$2:$A${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("A"), column_index_from_string("A"))}': {
        'H'},
    f'=\'Datengrundlage adesso\'!$E$2:$E${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("E"), column_index_from_string("E"))}': {
        'Z'},
    f'=\'Datengrundlage adesso\'!$G$2:$G${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("G"), column_index_from_string("G"))}': {
        'X', 'AA', 'AB', 'O', 'AG', 'AH', 'AI', 'AJ'},
    f'=\'Datengrundlage adesso\'!$I$2:$I${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("I"), column_index_from_string("I"))}': {
        'T'},
    f'=Formulare!$A$2:$A${get_max_row(wb["Formulare"], column_index_from_string("A"), column_index_from_string("A"))}': {
        'U'},
    f'=\'Datengrundlage adesso\'!$K$2:$K${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("K"), column_index_from_string("K"))}': {
        'AD'},
    f'=Übersicht!$F$2:$F${get_max_row(wb["Übersicht"], column_index_from_string("F"), column_index_from_string("F"))}': {
        'F'}
}


# TODO:
#   - the datavalidation objects must be modified and can neither be overwritten nor deleted, so these bad boys must get            # DONE
#     a new multicellrange - max value                                                                                              # DONE
#   - then the formula1 parameter must be checked and if this doesnt work, it needs to be replaced with the new values              # DONE
#   - further the Style of the cells must be expanded to the last row -> formatting rules?
#   - then the program should work fine for one BTT file, if there are more files at once, what happens with duplicates?
#   - also what should be done if there already is an existing consolidated file?
#   - and should the final concatenated file contain columns for the date and the source BTT file?
#   - there also seems to be a problem with the new file when opening


# this actually works and clears the dataValidations
wb["BTT"].data_validations = openpyxl.worksheet.datavalidation.DataValidationList()
add_dv(wb["BTT"])
wb.save('output.xlsx')
