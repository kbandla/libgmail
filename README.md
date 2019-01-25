# libgmail - Gmail IMAP Interface in Python
libgmail is a minimal API to access Gmail using python via Gmail's IMAP interface.
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

### Use Gmail's search operators
```
...
params = {'from': 'billg@microsoft.com', 'newer_than': '14d'}
emails = gmail.advanced_search(**params)
...
```

[Search operators you can use with Gmail](https://support.google.com/mail/answer/7190?hl=en)

## License
See the LICENSE file

## Requires
* Python 3 or later
