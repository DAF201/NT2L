import core.EXCEL
import core.TOOL


def is_feedfile(workbook: core.EXCEL.ExcelWorkBook) -> bool:
    """check if a workbook is feedfile"""
    # feedfile only has one sheet
    if workbook.sheets_count > 1:
        return False
    sheet = workbook[1]
    # check if the workbook has a colum SERIAL_NUMBER, if yes then check cell at 2 is empty or not
    try:
        if sheet[2, sheet.search("SERIAL_NUMBER").col].value:
            return False
    except:
        return False
    return True


def add_feedfile_sheets(workbook: core.EXCEL.ExcelWorkBook) -> None:
    """add the front and back sheets to feedfile"""
    if is_feedfile(workbook):
        workbook.add_sheet(1, "BEFORE", "Sheet 2")
        workbook.add_sheet(2, "AFTER", "Sheet 1")
        workbook.move_sheet(3, 2, "BEFORE")
        workbook.save()


def get_gr_value(workbook: core.EXCEL.ExcelWorkBook, key: str, range="") -> str | float:
    """get a value of the gr file, like BUILD for PB, ACTIVITY_QTY for qty..."""
    sheet: core.EXCEL.ExcelSheet
    if is_feedfile(workbook):
        sheet = workbook[1]
    else:
        sheet = workbook[2]
    try:
        return sheet[2, sheet.search(key, range).col].value
    except Exception as e:
        core.TOOL.alert(e)
        return None


def fill_gr_value(workbook: core.EXCEL.ExcelWorkBook, key: str, value: str, sn_size=None) -> None:
    """fill a column by a value, need the title of the column to find it"""
    if sn_size is None:
        sn_size = int(get_gr_value(workbook, "ACTIVITY_QTY"))
    sheet: core.EXCEL.ExcelSheet
    if is_feedfile(workbook):
        sheet = workbook[1]
    else:
        sheet = workbook[2]
    # move active cell to (1,1) to search from top left
    sheet[1, 1].select()
    target_col = sheet.search(key).col
    for i in range(2, sn_size+2):
        sheet[i, target_col].value = value
    workbook.save()


def fill_gr_sn(workbook: core.EXCEL.ExcelWorkBook, sn_start: int | str):
    """fill the serial number column, need to known the starting SN of an unit"""
    sn_start = int(sn_start)
    sheet: core.EXCEL.ExcelSheet
    if is_feedfile(workbook):
        sheet = workbook[1]
    else:
        sheet = workbook[2]
    sn_col = sheet.search("SERIAL_NUMBER").col
    sn_size = int(get_gr_value(workbook, "ACTIVITY_QTY"))
    for i in range(2, sn_size+2):
        sheet[i, sn_col].value = sn_start
        sheet[i, sn_col].text_color = 16711680
        sheet[1, sn_col].number_format = "0"
        sn_start += 1
    sheet.range(
        f"{core.EXCEL.col_num_to_letter(sn_col)}:{core.EXCEL.col_num_to_letter(sn_col)}").NumberFormat = "0"
    workbook.save()


def find_gr_next_empty_row(workbook: core.EXCEL.ExcelWorkBook, sheet_index: int) -> int:
    """find the next empty row by checking if the col A is empty, return the index of first empty row"""
    sheet = workbook[sheet_index].sheet
    last_cell = sheet.Cells(sheet.Rows.Count, 1).End(-4162)
    next_row = last_cell.Row + \
        1 if last_cell.Value not in (None, "") else last_cell.Row
    return next_row


def find_gr_cols_size(workbook: core.EXCEL.ExcelWorkBook) -> int:
    """find the number of columns in a sheet"""
    sheet = workbook[1]
    size = 1
    while sheet[1, size].value is not None:
        size += 1
    return size-1


def copy_gr_row(workbook: core.EXCEL.ExcelWorkBook, src_row: int, dst_row: int, src_sheet_index=2, target_sheet_index=1) -> None:
    """copy a row and paste to another one"""
    data_sheet = workbook[src_sheet_index]
    target_sheet = workbook[target_sheet_index]
    data_sheet.copy_and_paste(
        f"{src_row}:{src_row}", target_sheet, f"{dst_row}:{dst_row}")
    workbook.save()


def expand_gr_filter(workbook: core.EXCEL.ExcelWorkBook, sheet_index=2) -> None:
    """expand filtered rows"""
    sheet = workbook[sheet_index]
    sheet.show_all_data()
    workbook.save()


def remove_gr_fillter_row(workbook: core.EXCEL.ExcelWorkBook, sheet_index=1) -> None:
    """remove a column"""
    sheet = workbook[sheet_index]
    cell = sheet.search("=")
    if cell:
        col = core.EXCEL.col_num_to_letter(cell.col)
        sheet.clear(f"{col}:{col}")
    workbook.save()
