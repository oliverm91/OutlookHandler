### Only tested in few artificial cases. Never in real processes

# OutlookHandler

Wrapper of win32com library for easier use.

## Install instructions:
`python setup.py bdist_wheel`

Then install whl file in bdist folder with pip as

`pip install bdist/whlfilename.whl`

## Examples

### Find mail
```python
# Get Dict of folderName: List[ReceivedMail, ReceivedMail, ...]
hndlr = OutlookHandler('NameOfOutlookRootFolder') # Sometimes is own mail like x@x.com
received_mails_dict = hndlr.get_emails_by_subject(subject, exact_date=date.today(), search_in_inbox=True)

#Iter
for received_mails in received_mails_dict.values():
    for received_mail in received_mails:
        print(f'Replying to {received_mail.sender}')
        received_mail.reply_all('This is the reply', extra_copy_recipients=['x@x.com'], attachment_paths=[os.path.join(path, 'test.txt')])
```
### Send mail
```python
nm = NewMail(['x@x.com', 'y@x.com'], subject=subject, body='ABC', attachment_path=os.path.join(path, 'test1.txt'))
nm.add_recipient('z@x.com')
nm.add_copy_recipient('w@x.com')
nm.add_attachment_path(os.path.join(path, 'test2.txt'))
nm.send()
```
