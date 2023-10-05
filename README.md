### Only tested in few artificial cases. Never in real processes

# easy_outlook

Wrapper of win32com library for easier use.

## Install instructions:
If git is installed: `pip install git+https://github.com/oliverm91/easy_outlook.git@main`

Else, build a `whl` with `python setup.py bdist_wheel`. The `whl` file will be in a new generated dist/ folder. Then, install the `whl` file in dist/ folder with pip as `pip install bdist/whlfilename.whl`

## Examples

### Find, download attachment and reply mail
```python

outlook_root_folder_must_contain = 'NameOfOutlookRootFolder' # Sometimes is own mail like x@x.com
hndlr = OutlookHandler(outlook_root_folder_must_contain)
found_mails_must_contain_in_subject = 'ABC Report'

# Get dict of folderName: list[ReceivedMail, ReceivedMail, ...]
received_mails_dict = hndlr.get_emails_by_subject(found_mails_must_contain_in_subject, exact_date=date.today(), search_in_inbox=True)

#Iter
for received_mails in received_mails_dict.values():
    for received_mail in received_mails:
        print('Downloading attachments')
        for attachment in received_mail.attachments:
            print(attachment.filename, attachment.size)
            attachment.save('some_path', 'filename.ext')
        print(f'Replying to {received_mail.sender}')
        received_mail.reply_all('This is the reply', extra_copy_recipients=['x@x.com'], attachment_paths=[os.path.join(path, 'test.txt')])

```
### Send mail with many recipients, CC and attachments
```python
nm = NewMail(['x@x.com', 'y@x.com'], subject=subject, body='ABC', attachment_path=os.path.join(path, 'test1.txt'))
nm.add_recipient('z@x.com')
nm.add_copy_recipient('w@x.com')
nm.add_attachment_path(os.path.join(path, 'test2.txt'))
nm.send()
```
