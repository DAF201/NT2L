import config.regex_reference
import core.REGEX
import core.TOOL
import functions.gr_table
import functions.target_table
import core.EXCEL
import config.config
import service.report_build
import os


def build_report(report_date=core.TOOL.get_today_date()) -> list[int] | None:
    """build a daily report"""
    core.TOOL.message(
        f"start building report for {report_date}")

    target_table_excel = functions.target_table.target_table()


    report_excel = core.EXCEL.Excel()
    report_excel_workbook = report_excel.open(os.path.abspath(
        config.config.global_config["Excel"]["report_table"]))
    report_excel_sheet = report_excel_workbook[1]
    target_table_sheet = target_table_excel.target_sheet
    current_row = target_table_excel.search_date_start(
        report_date)

    if current_row == 0:
        core.TOOL.alert("target table not yet up to date")
        return None

    report_excel_sheet.clear()

    attention_needed = []

    while target_table_sheet[current_row, 2].value is not None:
        core.TOOL.message(f"checking target table row: {current_row}")
        if target_table_excel.check_row_ready_for_report(current_row):
            report_row = functions.gr_table.find_gr_next_empty_row(
                report_excel_workbook, 1)

            row_data = target_table_excel.search_row_for_info(
                current_row)

            # FXSJ, just copy from target table to report
            if row_data["DN#"] == "FXSJ":
                core.TOOL.message(
                    f"start building report row for {row_data["PB#"]}")
                report_excel_sheet[report_row,
                                   1].value = report_date
                report_excel_sheet[report_row, 2].value = row_data["PB#"]
                report_excel_sheet[report_row, 3].value = row_data["WO"]
                report_excel_sheet[report_row, 4].value = row_data["PN#"]
                report_excel_sheet[report_row, 5].value = "NA"
                report_excel_sheet[report_row, 6].value = "NA"
                report_excel_sheet[report_row, 7].value = row_data["DN Qty"]
                report_excel_sheet[report_row, 8].value = "NA"
                report_excel_sheet[report_row, 9].value = "Yes"

                # has something in note, so GI to WO...
                if target_table_sheet[current_row, 18].value:
                    report_excel_sheet[report_row,
                                       10].value = target_table_sheet[current_row, 18].value
                # nothing, so regular Vendor Pool
                else:
                    report_excel_sheet[report_row,
                                       10].value = "FXSJ Pooled Vendor Stock"
                report_excel_sheet[report_row, 11].value = "NA"
                report_excel_sheet[report_row, 12].value = "NA"
                report_excel_sheet[report_row, 13].value = "Done"
                report_excel_sheet[report_row, 14].value = row_data["FX SO"]
                report_excel_sheet[report_row, 15].value = row_data["FX DN"]
                report_excel_sheet[report_row, 16].value = row_data["SFC scan"]
                report_excel_sheet[report_row,
                                   17].value = row_data["FX GR invoice #"]
                report_excel_sheet[report_row, 18].value = "Done"
                report_excel_sheet[report_row, 19].value = "NA"

                report_excel_workbook.save()

            else:
                # regular dn, need to read the POD
                if core.REGEX.re_compare(config.regex_reference.DN_REGEX, str(row_data["DN#"])[:-2]):
                    dn_data = service.report_build.extract_pod_data(
                        str(int(row_data["DN#"])))
                    # cannot find data or amount mismatch
                    if dn_data is None:
                        attention_needed.append((current_row, row_data["PB#"], row_data["DN Qty"], row_data["DN#"], row_data["Ship to"],
                                                row_data["Shipment Status"], row_data["FX GR invoice #"], row_data["Owner"]))
                        current_row += 1
                        continue

                    for recipient in dn_data["TRANSACTION"]:
                        core.TOOL.message(
                            f"starting building report row for recipient number {recipient[1]}")
                        report_excel_sheet[report_row,
                                           1].value = report_date
                        report_excel_sheet[report_row,
                                           2].value = row_data["PB#"]
                        report_excel_sheet[report_row,
                                           3].value = row_data["WO"]
                        report_excel_sheet[report_row,
                                           4].value = row_data["PN#"]
                        report_excel_sheet[report_row, 5].value = recipient[1]
                        report_excel_sheet[report_row, 6].value = recipient[0]
                        report_excel_sheet[report_row, 7].value = recipient[2]
                        report_excel_sheet[report_row,
                                           8].value = row_data["DN#"]
                        report_excel_sheet[report_row, 9].value = "NA"
                        report_excel_sheet[report_row,
                                           10].value = row_data["Shipment Status"]
                        report_excel_sheet[report_row,
                                           11].value = row_data["Carrier"]
                        if row_data["Ship to"].lower() in ["sc", "zanker", "dallas", "austin", "nj", "nc", "ma"]:
                            report_excel_sheet[report_row,
                                               12].value = "NA"
                        else:
                            report_excel_sheet[report_row,
                                               12].value = "TRACKING# NEEDED"
                        report_excel_sheet[report_row,
                                           13].value = "Done"
                        report_excel_sheet[report_row,
                                           14].value = row_data["FX SO"]
                        report_excel_sheet[report_row,
                                           15].value = row_data["FX DN"]
                        report_excel_sheet[report_row,
                                           16].value = row_data["SFC scan"]
                        report_excel_sheet[report_row,
                                           17].value = row_data["FX GR invoice #"]
                        report_excel_sheet[report_row, 18].value = "Done"
                        report_excel_sheet[report_row, 19].value = "NA"
                        report_excel_workbook.save()
                else:
                    attention_needed.append((current_row, row_data["PB#"], row_data["DN Qty"], row_data["DN#"], row_data["Ship to"],
                                             row_data["Shipment Status"], row_data["FX GR invoice #"], row_data["Owner"]))
                    current_row += 1
                    continue
        current_row += 1
    report_excel.quit()
    return attention_needed if attention_needed != [] else None
