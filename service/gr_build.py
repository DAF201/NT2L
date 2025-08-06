import core.EXCEL
import core.TOOL
import functions.gr_table

# TODO: add alert after scanned an already GR done unit


def feedfile_to_gr_file(workbook: core.EXCEL.ExcelWorkBook, sn_start: int, status="PASS") -> None:
    """convert a feedfile to a GR file ready to be GR"""
    if sn_start == None:
        return
    if not functions.gr_table.is_feedfile(workbook):
        raise RuntimeError("This function can only be applied to FeedFile")
    functions.gr_table.add_feedfile_sheets(workbook)
    functions.gr_table.fill_gr_sn(workbook, sn_start)
    functions.gr_table.fill_gr_value(workbook, "PASS_FAIL_SCRAP", status)


def gr_file_building(workbook: core.EXCEL.ExcelWorkBook, sn_list: list[int], status="PASS") -> None:
    """making the gr file ready for SAP to use"""
    if functions.gr_table.is_feedfile(workbook):
        raise RuntimeError("This function can only be applied to GR file")

    gr_sheet = workbook[1]
    data_sheet = workbook[2]
    email_sheet = workbook[3]

    data_sheet.show_all_data()

    gr_sheet.clear()
    email_sheet.clear()
    functions.gr_table.copy_gr_row(workbook, 1, 1, 2, 1)
    email_sheet[1, 1].value = "SN"
    email_sheet[1, 2].value = "STT"

    sn_col = data_sheet.search("SERIAL_NUMBER").col
    sn_col_c = core.EXCEL.col_num_to_letter(sn_col)

    attention_list = []

    sn_list.sort()
    for sn in sn_list:
        data_sn_cell = data_sheet.search(sn, f"{sn_col_c}:{sn_col_c}")
        if data_sn_cell == None:
            core.TOOL.alert_and_beep("SN not found in this file")
        data_sheet_sn_row = data_sn_cell.row

        # GR done before
        if data_sheet[data_sheet_sn_row, sn_col].text_color == 255:
            attention_list.append(sn)
            continue

        data_sheet[data_sheet_sn_row, sn_col].text_color = 255

        email_sheet_next_empty_row = functions.gr_table.find_gr_next_empty_row(
            workbook, 3)
        gr_sheet_next_empty_row = functions.gr_table.find_gr_next_empty_row(
            workbook, 1)

        email_sheet[email_sheet_next_empty_row, 1].value = sn
        email_sheet[email_sheet_next_empty_row, 2].value = status

        data_sheet[data_sheet_sn_row, data_sheet.search(
            "PASS_FAIL_SCRAP").col].value = status

        functions.gr_table.copy_gr_row(
            workbook, data_sheet_sn_row, gr_sheet_next_empty_row, 2, 1)

    if attention_list != []:
        core.TOOL.alert_and_beep(
            f"Following SN has been GRed before and got skipped:\n{'\n'.join(attention_list)}")

    functions.gr_table.remove_gr_fillter_row(workbook, 1)
    workbook.save()


def gr_rename_data_extract(workbook: core.EXCEL.ExcelWorkBook) -> list[str]:
    """get the PB and Amount info of a gr sheet"""
    if functions.gr_table.is_feedfile(workbook):
        raise RuntimeError("This function can only be applied to GR file")
    pb = functions.gr_table.get_gr_value(workbook, "BUILD", "A:F")
    amount = functions.gr_table.find_gr_next_empty_row(workbook, 1)-2
    return pb, amount
