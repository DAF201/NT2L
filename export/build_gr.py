import functions.gr_table
import service.gr_build
import core.TOOL
import core.EXCEL
import service.gr_scan
import os


def build_gr() -> None:
    """API for build gr file from either gr already existed or feedfile"""
    file = core.TOOL.select_file("select a gr file or feedfile", [
                                 ("GR OR FEEDFILE", "*.xlsx")])
    if file == "":
        return
    gr_excel_app = core.EXCEL.Excel()
    gr_excel_workbook = gr_excel_app.open(file)

    range_allocation_sn = ""

    if functions.gr_table.is_feedfile(gr_excel_workbook):
        core.TOOL.message("scan a sn in this work order to start building")
        range_allocation_sn = service.gr_scan.scan_one_sn()
        service.gr_build.feedfile_to_gr_file(
            gr_excel_workbook, service.gr_scan.get_sn_range_from_sn(range_allocation_sn)[0])
        core.TOOL.message(
            "Feedfile convertion complete, GR file generated.")
        core.TOOL.message("starting GR scan")
    core.TOOL.focus_console()
    core.TOOL.message(
        "please enter status of those units (Default: 'PASS' by pressing Enter directly)")
    status = input()
    if status == "quit":
        gr_excel_app.quit()
        return
    if status == "":
        status = "PASS"
    core.TOOL.message(
        "please start scanning serial numbers. Press 'Enter' when finish scanning")
    sn_scan_data = service.gr_scan.scan_sn()
    if range_allocation_sn != "" and range_allocation_sn not in sn_scan_data:
        sn_scan_data.append(range_allocation_sn)
        sn_scan_data.sort()
    if sn_scan_data is None:
        gr_excel_app.quit()
        return
    service.gr_build.gr_file_building(gr_excel_workbook, sn_scan_data, status)
    pb, amount = service.gr_build.gr_rename_data_extract(gr_excel_workbook)
    gr_excel_app.quit()
    base_dir = os.path.abspath(file).replace(os.path.basename(file), "")
    gr_excel_app.rename(
        f"{base_dir}\\{core.TOOL.get_today_date_pad2()} {pb}_{amount}x_{core.TOOL.get_invoice()}.xlsx")
