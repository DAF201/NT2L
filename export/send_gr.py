import core.OUTLOOK
import core.TOOL
import core.EXCEL
import config.config
import service.email_build
import os


def send_gr() -> None:
    """API for send gr file to NV"""
    attachment = core.TOOL.select_file(
        "select a gr file", [("GR FILE", "*.xlsx")])

    # DN always need an attachment
    if attachment == "":
        return

    gr_excel_app = core.EXCEL.Excel()
    gr_excel_workbook = gr_excel_app.open(attachment)

    file_data = os.path.basename(attachment).split("_")
    amount = file_data[1]
    pb = file_data[0][-8:]
    invoice = file_data[2][:-5]

    subject = str(config.config.global_config["DN"]["DN_SUBJECT"]).format(
        amount, pb, invoice)

    to = "fangzhou.ye@fii-na.com"
    cc = "fangzhou.ye@fii-na.com"
    # to = config.config.global_config["DN"]["DN_mail_To"]
    # cc = config.config.global_config["DN"]["DN_mail_CC"]

    html_table = service.email_build.build_html_table(
        service.email_build.get_sn_and_status(gr_excel_workbook))

    print("Please enter note to NV PM, press 'ENTER' when finished")
    note = input()
    if note != "":
        note = "<br><br>NOTE: " + note

    html_body = str(config.config.global_config["DN"]["DN_HTML_BODY"]).format(
        note, amount, pb, invoice, html_table, config.config.global_config["Outlook"]["USER"])

    core.OUTLOOK.global_outlook_app.send_email(
        subject=subject, to=to, cc=cc, HTMLBody=html_body, attachments=attachment)

    gr_excel_app.quit()
