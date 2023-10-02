from typing import Dict, List, Union
import win32com.client
from datetime import datetime, date
import os

class NewMail:
    def __init__(self, recipient: Union[str, List[str]], copy_recipient: Union[str, List[str]]=None, subject: str="", body: str="", html_body: str="", attachment_path: Union[str, List[str]]=None):
        if isinstance(recipient, str):
            self._recipient = [recipient]            
        else:
            self._recipient = [v for v in recipient]
        if isinstance(copy_recipient, str):
            self._copy_recipient = [copy_recipient]            
        else:
            if copy_recipient is None:
                copy_recipient = []
            self._copy_recipient = [v for v in copy_recipient]
        self._subject = subject
        self._body=body
        self._html_body=html_body
        if isinstance(attachment_path, str):
            self._attachment_path = [attachment_path]            
        else:
            self._attachment_path = [v for v in attachment_path]
        self._mail=None
        self.set_mail_obj()

    @property
    def recipient(self) -> List[str]:
        return self._recipient

    @recipient.setter
    def recipient(self, value: Union[str, List[str]]):
        if isinstance(value, str):
            self._recipient = [value]            
        else:
            self._recipient = [v for v in value]
        self.set_mail_obj()

    def add_recipient(self, new_recipient: str) -> None:
        r = self.recipient.copy()
        r.append(new_recipient)
        self.recipient = r

    @property
    def copy_recipient(self) -> List[str]:
        return self._copy_recipient

    @copy_recipient.setter
    def copy_recipient(self, value: Union[str, List[str]]):
        if isinstance(value, str):
            self._copy_recipient = [value]            
        else:
            self._copy_recipient = [v for v in value]
        self.set_mail_obj()

    def add_copy_recipient(self, new_copy_recipient: str) -> None:
        cr = self.copy_recipient.copy()
        cr.append(new_copy_recipient)
        self.copy_recipient = cr

    @property
    def subject(self) -> str:
        return self._subject

    @subject.setter
    def subject(self, value: str):
        self._subject = value
        self.set_mail_obj()

    @property
    def body(self) -> str:
        return self._body

    @body.setter
    def body(self, value: str):
        self._body = value
        self.set_mail_obj()

    @property
    def html_body(self) -> str:
        return self._html_body

    @html_body.setter
    def html_body(self, value: str):
        self._html_body = value
        self.set_mail_obj()

    @property
    def attachment_path(self) -> List[str]:
        return self._attachment_path

    @attachment_path.setter
    def attachment_path(self, value: Union[str, List[str]]):
        if isinstance(value, str):
            self._attachment_path = [value]            
        else:
            self._attachment_path = [v for v in value]
        self.set_mail_obj()

    def add_attachment_path(self, attachment_path: str) -> None:
        ap = self.attachment_path.copy()
        ap.append(attachment_path)
        self.attachment_path = ap

    def remove_duplicate_recipients(self) -> None:
        recipients = [r for r in self._mail.Recipients if r.Type==1]
        cc_recipients = [r for r in self._mail.Recipients if r.Type==2]
        recipient_addresses = []
        cc_recipient_addresses = []
        for r in recipients:
            address = str(r.AddressEntry)
            if address in recipient_addresses:
                self._mail.Recipients.Remove(r.Index)
            recipient_addresses.append(address)
        for cc_r in cc_recipients:
            address = str(cc_r.AddressEntry)
            if address in cc_recipient_addresses:
                self._mail.Recipients.Remove(cc_r.Index)
            cc_recipient_addresses.append(address)

    def set_mail_obj(self) -> None:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        to_recipients = [self.recipient] if isinstance(self.recipient, str) else self.recipient

        for rec in to_recipients:
            r = mail.Recipients.Add(rec)
            r.Type = 1 # Declare recipient in To
        if self.copy_recipient is not None:
            cc_recipients = [self.copy_recipient] if isinstance(self.copy_recipient, str) else self.copy_recipient
            for rec in cc_recipients:
                r = mail.Recipients.Add(rec)
                r.Type = 2 # Declare recipient in CC
        if self.subject != "":
            mail.Subject = self.subject
        if self.html_body != "":
            mail.HTMLBody = self.html_body
        elif self.body != "":
            mail.Body = self.body
        
        if self.attachment_path is not None:
            attachment_paths = [self.attachment_path] if isinstance(self.attachment_path, str) else self.attachment_path
            for ap in attachment_paths:
                mail.Attachments.Add(ap)
        
        self._mail=mail
        self.remove_duplicate_recipients()
    
    def send(self) -> None:
        self._mail.Send()


class ReceivedMailAttachment:
    def __init__(self, pyWin32AttachmentObj) -> None:
        self.pywin32attachment = pyWin32AttachmentObj
        self.filename: str = self.pywin32attachment.filename
        self.size: int = self.pywin32attachment.size #Size in bytes
    
    def save(self, save_dir: str, save_filename: str) -> None:
        self.pywin32attachment.SaveAsFile(os.path.join(save_dir, save_filename))


class ReceivedMail:
    def __init__(self, pyWin32MailObj):
        self.pywin32mail = pyWin32MailObj
        self.datetime: datetime = datetime(self.pywin32mail.ReceivedTime.year, self.pywin32mail.ReceivedTime.month, self.pywin32mail.ReceivedTime.day,
                                 self.pywin32mail.ReceivedTime.hour, self.pywin32mail.ReceivedTime.minute, self.pywin32mail.ReceivedTime.second)
        self.date: date = self.datetime.date()
        self.subject: str = self.pywin32mail.subject
        self.sender = str(self.pywin32mail.Sender)
        self.body: str = self.pywin32mail.body
        self.html_body: str = self.pywin32mail.htmlbody
        pywin32_attachments_lst = list(self.pywin32mail.attachments)
        self.attachments: List[ReceivedMailAttachment] = [ReceivedMailAttachment(pywin32_att) for pywin32_att in pywin32_attachments_lst if pywin32_att.type == 1]
        self.has_attachments: bool = len(self.attachments) > 0
    
    def __str__(self) -> str:
        return f'<ReceivedMail obj: {self.subject[:10]}..., from: {self.sender}, sent on: {self.date.strftime("%Y-%m-%d")}'
    
    def __repr__(self) -> str:
        return self.__str__()
    
    def reply_all(self, body: str="", html_body: str="", extra_recipients: List[str]=None, extra_copy_recipients: List[str]=None, attachment_paths: List[str]=None) -> None:
        reply = self.pywin32mail.ReplyAll()
        if html_body != "":
            reply.HTMLBody = html_body + reply.HTMLBody
        elif body != "":
            reply.Body = body + reply.Body
        if extra_recipients is not None:
            for er in extra_recipients:
                r=reply.Recipients.Add(er)
                r.Type=1
        if extra_copy_recipients is not None:
            for ecr in extra_copy_recipients:
                r=reply.Recipients.Add(ecr)
                r.Type=2
        if attachment_paths is not None:
            for ap in attachment_paths:
                reply.Attachments.Add(ap)
        reply.send


class OutlookHandler:
    def __init__(self, root_folder_name_contain: str, inbox_name: str='Bandeja de entrada', max_tries_found_root_folder: int=30) -> None:
        self.outlook_app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.root_folder_name_contain = root_folder_name_contain
        self.root_folder = self.get_root_folder()
        self.inbox_folder = [f for f in self.root_folder.Folders if f.name==inbox_name][0]
        self._mtfrf = max_tries_found_root_folder

    def get_root_folder(self) -> str:
        counter = 0
        while counter < self._mtfrf:
            folder = self.outlook_app.Folders.Item(counter)
            if self.root_folder_name_contain in folder.name:
                return folder
            counter += 1
            
        raise LookupError(f'Root folder not found. No folder name containing {self.root_folder_name_contain} was found.')

    def _search_emails_by_subject_recursive(self, folder, subject_contains, folder_mails_dict, min_date: date=None, max_date: date=None, exact_date: date=None, folders: List[str]=None, search_in_inbox: bool=False):
        search_filter = f"@SQL=urn:schemas:httpmail:subject LIKE '%{subject_contains}%'"
        if exact_date is not None:
            formatted_exact_date = exact_date.strftime('%Y-%m-%d')
            search_filter += f" AND (urn:schemas:httpmail:datereceived >= '{formatted_exact_date} 00:00:00' AND urn:schemas:httpmail:datereceived <= '{formatted_exact_date} 23:59:59')"
        else:
            if min_date is not None:
                formatted_start_date = min_date.strftime('%Y-%m-%d')
                search_filter += f" AND (urn:schemas:httpmail:datereceived >= '{formatted_start_date}')"
            if max_date is not None:
                formatted_end_date = max_date.strftime('%Y-%m-%d')
                search_filter += f" AND (urn:schemas:httpmail:datereceived <= '{formatted_end_date}')"
        
        filtered_emails = folder.Items.Restrict(search_filter)

        if filtered_emails.count > 0:
            pywin_mails_lst = list(filtered_emails)
            mail_lst = [ReceivedMail(pywin_mail) for pywin_mail in pywin_mails_lst]
            folder_mails_dict[folder.name] = mail_lst

        # Recursively search in subfolders.
        if search_in_inbox:
            folders = [self.inbox_folder.name.lower()]
        if folder.Folders.Count > 0:
            for subfolder in folder.Folders:
                if folders is not None:
                    if subfolder.name.lower() not in folders:
                        continue
                self._search_emails_by_subject_recursive(subfolder, subject_contains, folder_mails_dict, min_date=min_date, max_date=max_date, exact_date=exact_date, folders=folders)

    def get_emails_by_subject(self, subject_contains: str, min_date: date=None, max_date: date=None, exact_date: date=None, folders: List[str]=None, search_in_inbox: bool=False) -> Dict[str, List[ReceivedMail]]:
        # Start the recursive search from the root folder.
        folder_mails_dict = {}
        if search_in_inbox:
            folders = None
        if folders is not None:
            folders = [folder.lower() for folder in folders if folder is not None]
            if len(folders) == 0:
                folders = None
        self._search_emails_by_subject_recursive(self.root_folder, subject_contains, folder_mails_dict, min_date=min_date, max_date=max_date, exact_date=exact_date, folders=folders, search_in_inbox=search_in_inbox)
        return folder_mails_dict