# libgmail - Gmail IMAP Interface in Python
libgmail is a simple API to access Gmail using python via Gmail's IMAP interface.
It is work in progress, based on my needs. Patches are welcome. 

## Examples
Here are some examples of libgmail in action

### Get Attachments for the past 5 days
```python
from libgmail import Gmail
mailbox = Gmail(username, password)
attachments = mailbox.getAttachmentsForDays(5)
for attachment in attachments:
    # write each attachment to /tmp
    with open('/tmp/%s'%attachment.filename,'w') as f:
        f.write(attachment.data)
```

### Get emails from a user
```python
from libgmail import Gmail
x=Gmail(username, password)
emails = x.search('from: billg@microsoft.com')
```

## Author
Kiran Bandla

## License
See the LICENSE file

## Requires
* Python 2.6 or later
