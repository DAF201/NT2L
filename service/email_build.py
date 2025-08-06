import typing
import core.EXCEL


def get_sn_and_status(workbook: core.EXCEL.ExcelWorkBook) -> list[typing.Union[typing.Tuple[str, str], typing.Tuple[str, str, str]]]:
    """read the GR sheet, get all the data"""
    GR_data = []
    email_sheet = workbook[3]
    current_row = 2
    while True:
        if email_sheet[current_row, 1].value is None:
            break
        A_data = email_sheet[current_row, 1].value
        B_data = email_sheet[current_row, 2].value
        C_data = email_sheet[current_row, 3].value
        if C_data:
            GR_data.append((str(int(A_data)), B_data, str(int(C_data))))
        else:
            GR_data.append((str(int(A_data)), B_data))
        current_row += 1
    return GR_data


def build_html_table(data: list) -> str:
    """The email need to insert HTML to show a table"""

    table_style = "border-collapse: collapse;"
    cell_style = "border: 1px solid black; padding: 4px; text-align: center;"

    rows = []

    has_config = (len(data[0]) == 3)

    if not has_config:
        header = f"<tr><th style='{cell_style}'>SN</th><th style='{cell_style}'>Status</th></tr>"
        for a, b in data:
            rows.append(
                f"<tr><td style='{cell_style}'>{a}</td><td style='{cell_style}'>{b}</td></tr>")

    else:
        header = f"<tr><th style='{cell_style}'>SN</th><th style='{cell_style}'>Status</th><th style='{cell_style}'>Config</th></tr>"
        for a, b, c in data:
            rows.append(
                f"<tr><td style='{cell_style}'>{a}</td><td style='{cell_style}'>{b}</td><td style='{cell_style}'>{c}</td></tr>")

    return f"""
    <html>
        <body>
        <table style="{table_style}">
            {header}
            {"".join(rows)}
        </table>
        </body>
    </html>
    """
