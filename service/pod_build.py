import core.OUTLOOK
import core.TOOL
import os


def download_pods(saving_directory: str) -> list[str]:
    """download the pdf attachement from outlook"""
    file_paths = []
    new_emails = core.OUTLOOK.global_outlook_app.get_emails(10)
    core.TOOL.message("collecting POD from Outlook")
    for mail in new_emails:
        if mail.Class != 43:
            continue
        if mail.attachments.Count == 1 and mail.subject == "Scanned image from MX-C301W":
            attachment = mail.attachments.Item(1)
            if attachment.FileName.lower().endswith(".pdf"):
                saving_path = os.path.join(
                    saving_directory, attachment.FileName)
                mail.attachments.Item(1).SaveAsFile(saving_path)
                file_paths.append(saving_path)
    core.TOOL.message(f"collection complete, {len(file_paths)} PODs found")
    return file_paths
