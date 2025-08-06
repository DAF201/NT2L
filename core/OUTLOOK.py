import os
import win32com.client as win32


class OutlookEmail:
    """outlook email, contains some information"""

    def __init__(self, email):
        self.email = email

    @property
    def subject(self):
        return self.email.Subject

    @property
    def body(self):
        return self.email.Body

    @property
    def HTML_body(self):
        return self.email.HTMLBody

    @property
    def sender_name(self):
        return self.email.SenderName

    @property
    def sender_email_address(self):
        return self.email.SenderEmailAddress

    @property
    def to(self):
        return self.email.TO

    @property
    def cc(self):
        return self.email.CC

    @property
    def sent_on(self):
        return self.email.SentOn

    @property
    def received_on(self):
        return self.email.ReceivedTime

    @property
    def attachments(self):
        return self.email.Attachments

    @property
    def Class(self):
        return self.email.Class

    def save_attachment(self, save_dir: str) -> None:
        """save attachment of this email to given folder"""
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        for i in range(1, self.attachments.Count + 1):
            attachment = self.attachments.Item(i)
            file_path = os.path.join(save_dir, attachment.FileName)
            attachment.SaveAsFile(file_path)


class Outlook:
    """outlook app, can read and send email"""

    def __init__(self):
        self.app = win32.Dispatch("Outlook.Application")
        self.inbox = self.app.GetNamespace("MAPI").GetDefaultFolder(6).Items

    def get_emails(self, num_of_mails=20) -> list[OutlookEmail]:
        """get the newest {} emails"""
        res = []
        self.inbox.Sort("[ReceivedTime]", True)
        email = self.inbox.GetFirst()
        while email and len(res) < num_of_mails:
            if email.Class == 43:
                res.append(OutlookEmail(email))
            email = self.inbox.GetNext()
        return res

    def send_email(self, subject: str,  to: str, cc=None, body="", HTMLBody=None, attachments=None) -> None:
        """send a new email"""
        out_going_mail = self.app.CreateItem(0)

        if isinstance(to, (list, tuple)):
            out_going_mail.To = ";".join(to)
        else:
            out_going_mail.To = to

        if cc:
            if isinstance(cc, (list, tuple)):
                out_going_mail.CC = ";".join(cc)
            else:
                out_going_mail.CC = cc

        out_going_mail.Subject = subject

        if HTMLBody:
            out_going_mail.HTMLBody = HTMLBody
        else:
            out_going_mail.Body = body

        if attachments:
            if isinstance(attachments, list) or isinstance(attachments, tuple):
                for file in attachments:
                    if isinstance(file, str):
                        out_going_mail.Attachments.Add(file)
            elif isinstance(attachments, str):
                out_going_mail.Attachments.Add(attachments)

        try:
            out_going_mail.Send()
        except Exception as e:
            print("Send email failed:", e)


global_outlook_app = Outlook()
