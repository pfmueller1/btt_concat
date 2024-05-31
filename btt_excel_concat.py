import tkinter
from tkinter import filedialog

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.packaging.manifest import Override
from openpyxl.reader.excel import load_workbook
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation, DataValidationList
from openpyxl.styles import PatternFill, Color
from openpyxl.formatting.rule import Rule
import xxhash
from openpyxl.worksheet.worksheet import Worksheet


class CustomWorksheet(Worksheet):
    def delete_rows(self, idx, min_col=None, max_col=None):
        if min_col and max_col:
            remainder = range(max(get_max_row(self, min_col, max_col), idx), min(idx+1, get_max_row(self, min_col, max_col))+1)
            max_col += 1
        else:
            remainder = range(max(self.max_row, idx), min(idx+1, self.max_row)+1)
            min_col = self.min_column
            max_col = self.max_column+1
        for row in remainder:
            for col in range(min_col, max_col):
                if (row, col) in self._cells:
                    del self._cells[row, col]
        self._current_row = self.max_row
        if not self._cells:
            self._current_row = 0


openpyxl.worksheet.worksheet.Worksheet.delete_rows = CustomWorksheet.delete_rows


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
            except Exception:
                continue
            count_col += 1
        count_row += 1
    del_dupes(sheet_receiving, columns=range(start_col, end_col + 1))
    print(sheet_receiving)


def add_dv(sheet, dv_list):
    max_row = sheet.max_row

    for formula, cols in dv_list.items():
        for col in cols:
            col_letter = col if col.isalpha() else get_column_letter(col)
            dv = DataValidation(
                type="list",
                formula1=f"{formula}",
                allow_blank=True,
                showDropDown=False,
                showInputMessage=True,
                showErrorMessage=True
            )
            range_str = f"${col_letter}$3:${col_letter}${max_row}"
            dv.add(range_str)
            sheet.add_data_validation(dv)


def add_cf(sheet, cf_list):
    red_fill = PatternFill(patternType=None,
                           fgColor=Color(rgb="000000",
                                         type="rgb"),
                           bgColor=Color(rgb="FFA7A7",
                                         type="rgb")
                           )

    dxf = DifferentialStyle(fill=red_fill)

    for formula, cols in cf_list.items():
        for col in cols:
            col_letter = col if col.isalpha() else get_column_letter(col)
            sheet.conditional_formatting.add(f"{col_letter}3:{col_letter}{sheet.max_row}",
                                             Rule(type="expression",
                                                  dxf=dxf,
                                                  formula=[f"{formula.replace('~', col_letter)}"])
                                             )


def hash_row(row):
    row_str = ''.join([str(cell) for cell in row])
    return xxhash.xxh64(row_str).hexdigest()


def del_dupes(sheet, columns=None):
    seen_hashes = {}
    del_list = []

    if columns:
        max_row = get_max_row(sheet, min(columns), max(columns))
    else:
        max_row = sheet.max_row

    for row_idx in range(2, max_row + 1):
        row_values = []
        for col in columns:
            cell = sheet.cell(row=row_idx, column=col)
            cell_value = cell.value
            row_values.append(cell_value)
        row_hash = hash_row(row_values)
        if row_hash in seen_hashes:
            del_list.append(row_idx)
        else:
            seen_hashes[row_hash] = row_idx

    for row_idx in reversed(del_list):
        if columns:
            sheet.delete_rows(row_idx, min_col=min(columns), max_col=max(columns))
        else:
            sheet.delete_rows(row_idx)


def main():
    root = tkinter.Tk()
    root.withdraw()
    wb = load_workbook('BTT_Template.xlsx')
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    all_tab_data = {}

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
                    {"start_col": 1, "start_row": 4, "end_col": 2, "end_row": get_max_row(wb_tmp[sheet_name], 1, 2)},  # [A:B]
                    {"start_col": 5, "start_row": 2, "end_col": 8, "end_row": get_max_row(wb_tmp[sheet_name], 5, 8)}  # [E:H]
                ]
            elif sheet_name == "BTT":
                tab_data[sheet_name] = [
                    {"start_col": 1, "start_row": 3, "end_col": wb_tmp[sheet_name].max_column,
                     "end_row": wb_tmp[sheet_name].max_row}
                ]
            elif sheet_name == "BPML":
                tab_data[sheet_name] = [
                    {"start_col": 1, "start_row": 2, "end_col": 4, "end_row": get_max_row(wb_tmp[sheet_name], 1, 4)},  # [A:D]
                    {"start_col": 6, "start_row": 2, "end_col": 10, "end_row": get_max_row(wb_tmp[sheet_name], 6, 10)}  # [F:J]
                ]
            elif sheet_name == "Transaktionen":
                tab_data[sheet_name] = [
                    {"start_col": 1, "start_row": 2, "end_col": 7, "end_row": get_max_row(wb_tmp[sheet_name], 1, 7)},  # [A:G]
                ]
            elif sheet_name == "Formulare":
                tab_data[sheet_name] = [
                    {"start_col": 1, "start_row": 2, "end_col": 3, "end_row": get_max_row(wb_tmp[sheet_name], 1, 3)},  # [A:G]
                ]
            elif sheet_name == "Schnittstellen":
                tab_data[sheet_name] = [
                    {"start_col": 1, "start_row": 2, "end_col": 6, "end_row": get_max_row(wb_tmp[sheet_name], 1, 6)},  # [A:F]
                    {"start_col": 8, "start_row": 2, "end_col": 10, "end_row": get_max_row(wb_tmp[sheet_name], 8, 10)}  # [F:J]
                ]
            elif sheet_name == "Datengrundlage adesso":
                tab_data[sheet_name] = [
                    {"start_col": 1, "start_row": 2, "end_col": 3, "end_row": get_max_row(wb_tmp[sheet_name], 1, 3)},  # [A:C]
                    {"start_col": 5, "start_row": 2, "end_col": 5, "end_row": get_max_row(wb_tmp[sheet_name], 5, 5)},  # [E]
                    {"start_col": 7, "start_row": 2, "end_col": 7, "end_row": get_max_row(wb_tmp[sheet_name], 7, 7)},  # [G]
                    {"start_col": 9, "start_row": 2, "end_col": 9, "end_row": get_max_row(wb_tmp[sheet_name], 9, 9)},  # [I]
                    {"start_col": 11, "start_row": 2, "end_col": 11, "end_row": get_max_row(wb_tmp[sheet_name], 11, 11)}  # [K]
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
                    sel_range = copy_range(d["start_col"], d["start_row"], d["end_col"], d["end_row"],
                                           wb_tmp[sheet_name])
                    paste_range(d["start_col"], get_max_row(wb[sheet_name], d["start_col"], d["end_col"]) + 1,
                                d["end_col"],
                                wb[sheet_name].max_row + d["end_row"], wb[sheet_name], sel_range)
            else:
                sel_range = copy_range(dim["start_col"] + 1, dim["start_row"] + 1, dim["end_col"], dim["end_row"],
                                       wb_tmp[sheet_name])
                paste_range(wb[sheet_name].min_column,
                            get_max_row(wb[sheet_name], dim["start_col"], dim["end_col"]) + 1,
                            wb[sheet_name].max_column,
                            wb[sheet_name].max_row + dim["end_row"], wb[sheet_name], sel_range)

            # expand table dimensions in template file
            if sheet_name == "Übersicht":
                wb[sheet_name].tables["Teilprojekte"].ref = f"$E$1:$H${get_max_row(wb[sheet_name], 5, 8)}"
            elif sheet_name == "BTT":
                wb[sheet_name].tables["BTT"].ref = f"$A$2:$AT${wb[sheet_name].max_row}"
            elif sheet_name == "BPML":
                wb[sheet_name].tables["Hauptprozesse"].ref = f"$A$1:$D${get_max_row(wb[sheet_name], 1, 4)}"
                wb[sheet_name].tables["BPML"].ref = f"$F$1:$J${get_max_row(wb[sheet_name], 6, 10)}"
            elif sheet_name == "Transaktionen":
                wb[sheet_name].tables["Transaktionen"].ref = f"$A$1:$G${get_max_row(wb[sheet_name], 1, 7)}"
            elif sheet_name == "Formulare":
                wb[sheet_name].tables["Formulare"].ref = f"$A$1:$C${get_max_row(wb[sheet_name], 1, 3)}"
            elif sheet_name == "Schnittstellen":
                wb[sheet_name].tables["Schnittstelle_Klarname"].ref = f"$H$1:$J${get_max_row(wb[sheet_name], 8, 10)}"
                wb[sheet_name].tables["Schnittstellen_technisch"].ref = f"$A$1:$F${get_max_row(wb[sheet_name], 1, 6)}"
            elif sheet_name == "Datengrundlage adesso":
                wb[sheet_name].tables["Module"].ref = f"$A$1:$C${get_max_row(wb[sheet_name], 1, 3)}"
                wb[sheet_name].tables["Prioritäten"].ref = f"$E$1:$E${get_max_row(wb[sheet_name], 5, 5)}"
                wb[sheet_name].tables["Vorhanden?"].ref = f"$G$1:$G${get_max_row(wb[sheet_name], 7, 7)}"
                wb[sheet_name].tables["Outputs"].ref = f"$I$1:$I${get_max_row(wb[sheet_name], 9, 9)}"
                wb[sheet_name].tables["Interfaces"].ref = f"$K$1:$K${get_max_row(wb[sheet_name], 11, 11)}"

    # distribute drop down lists
    dv_list = {
        f'BPML!$A$2:$A${get_max_row(wb["BPML"], column_index_from_string("A"), column_index_from_string("A"))}': {'B'},
        f'=BPML!$F$2:$F${get_max_row(wb["BPML"], column_index_from_string("F"), column_index_from_string("F"))}': {'C'},
        f'=Transaktionen!$A$2:$A${get_max_row(wb["Transaktionen"], column_index_from_string("A"), column_index_from_string("A"))}': {'I'},
        f'=Schnittstellen!$H$2:$H${get_max_row(wb["Schnittstellen"], column_index_from_string("H"), column_index_from_string("H"))}': {'R'},
        f'=\'Datengrundlage adesso\'!$A$2:$A${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("A"), column_index_from_string("A"))}': {'H'},
        f'=\'Datengrundlage adesso\'!$E$2:$E${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("E"), column_index_from_string("E"))}': {'Z'},
        f'=\'Datengrundlage adesso\'!$G$2:$G${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("G"), column_index_from_string("G"))}': {'X', 'AA', 'AB', 'O', 'AG', 'AH', 'AI', 'AJ'},
        f'=\'Datengrundlage adesso\'!$I$2:$I${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("I"), column_index_from_string("I"))}': {'T'},
        f'=Formulare!$A$2:$A${get_max_row(wb["Formulare"], column_index_from_string("A"), column_index_from_string("A"))}': {'U'},
        f'=\'Datengrundlage adesso\'!$K$2:$K${get_max_row(wb["Datengrundlage adesso"], column_index_from_string("K"), column_index_from_string("K"))}': {'AD'},
        f'=Übersicht!$F$2:$F${get_max_row(wb["Übersicht"], column_index_from_string("F"), column_index_from_string("F"))}': {'F'}
    }

    # distribute conditional formatting
    cf_list = {
        f"ISBLANK(B3)": {'B'},
        f'AND(ISBLANK(W3),OR(T3="Mail",T3="XML",T3="weiterer"))': {'W'},
        f'AND(ISBLANK(U3),T3="SAP-Formular")': {'U', 'V'},
        f'ISBLANK(~3)': {'H', 'I', 'D', 'O', 'T', 'X', 'Z'}
    }

    # TODO:
    #   - the datavalidation objects must be modified and can neither be overwritten nor deleted, so these bad boys must get            # DONE
    #     a new multicellrange - max value                                                                                              # DONE
    #   - then the formula1 parameter must be checked and if this doesnt work, it needs to be replaced with the new values              # DONE
    #   - further the Style of the cells must be expanded to the last row -> formatting rules?                                          # DONE
    #   - then the program should work fine for one BTT file, if there are more files at once, what happens with duplicates?            # DONE
    #   - also what should be done if there already is an existing consolidated file?   -> overwritten                                  # DONE
    #   - and should the final concatenated file contain columns for the date and the source BTT file?
    #   - there also seems to be a problem with the new file when opening                                                               # DONE
    #   - column AL? - aktivesTeilprojekt??
    #   - optimize performance!

    #del_dupes(wb["Datengrundlage adesso"], columns=[1, 2, 3])

    # modify data validations
    wb["BTT"].data_validations = DataValidationList()
    add_dv(wb["BTT"], dv_list)

    # modify conditional formatting
    wb.conditional_formatting = ConditionalFormattingList()
    add_cf(wb["BTT"], cf_list)

    wb.save('output.xlsx')
    wb.close()


if __name__ == "__main__":
    main()
