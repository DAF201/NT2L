import core.OUTLOOK
import core.TOOL
import config.config


def send_ITN(dn: str) -> None:
    """API for send ITN application"""
    attachment = core.TOOL.select_file(
        "select a SLI file", [("SLI", "*.pdf")])
    if attachment == "":
        return

    subject = str(config.config.global_config["ITN"]["ITN_SUBJECT"]).format(dn)
    body = str(config.config.global_config["ITN"]["ITN_BODY"]).format(
        dn, config.config.global_config["Outlook"]["USER"])
    to = config.config.global_config["ITN"]["ITN_mail_To"]
    cc = config.config.global_config["ITN"]["ITN_mail_CC"]
    core.OUTLOOK.global_outlook_app.send_email(
        subject=subject, to=to, cc=cc, body=body, attachments=attachment)
