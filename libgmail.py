#!/usr/bin/python
"""
 Gmail IMAP Interface in Python
 This is an interface to Google's IMAP extensions interface
 Kiran Bandla <kbandla@in2void.com>
 Fri Sep  7 18:23:09 EDT 2011

 References:
    *   http://docs.python.org/library/imaplib.html#imaplib.IMAP4_SSL
    *   https://developers.google.com/google-apps/gmail/imap_extensions
    *   https://support.google.com/mail/bin/answer.py?hl=en&answer=7190
    *   http://www.doughellmann.com/PyMOTW/imaplib/
    *   http://www.example-code.com/csharp/imap-search-critera.asp
    *   http://yuji.wordpress.com/2011/06/22/python-imaplib-imap-example-with-gmail/
    *   http://stackoverflow.com/questions/3283460/fetch-an-email-with-imaplib-but-do-not-mark-it-as-seen
    *   http://stackoverflow.com/questions/7930686/search-for-messages-with-attachments-with-gmail-imap
"""
import re
import imaplib
import email
import time
import logging
import datetime
from hashlib import md5, sha256

# Ignore the following attachments
IGNORE_LIST = []

__version__ = '0.4.0'
logger = logging.getLogger('libgmail')

SEARCH_KEYS = [
    'from',
    'to',
    'subject',
    'label',
    'list',
    'filename',
    'has',
    'in',
    'is',
    'cc',
    'bcc',
    'after',
    'before',
    'older',
    'newer',
    'older_than',
    'newer_than',
    'size',
    'larger',
    'smaller',
    'Rfc822msgid'
]


class GmailException(Exception):
    pass


class Attachment:
    def __init__(self, metadata={}):
        self.__dict__.update(metadata)
        self.__calchash()
    
    def __calchash(self):
        self.md5sum = md5(self.data).hexdigest()
        self.sha256 = sha256(self.data).hexdigest()

    def __str__(self):
        return self.__repr__()

    def __repr__(self):
        return self.filename


class Email:
    def __init__(self, msg_data, num):
        """
        @msg_data   : email message
        @num        : IMAP email id
        """
        self.msg = email.message_from_string(msg_data)
        self.num = num
        self.attachments = []
        self.subject = self.msg.get('subject')
        self.from_addr = self.msg.get('from')
        self.to_addr = self.msg.get('to')
        self.timestamp = None
        if self.msg.get('date'):
            self.timestamp = time.struct_time(email.utils.parsedate_tz(self.msg.get('date'))[:9])

    def __str__(self):
        return self.__repr__()

    def __repr__(self):
        return f'[{self.subject}]'

    def extractAttachments(self):
        """
        Extract attachments from an Email object
        appends to the list of Attachment objects property 'attachments'
        """
        msg = self.msg
        for part in msg.walk():
            if part.get('Content-Disposition'):
                if part.get('Content-Disposition').startswith('attach'):
                    # This looks like an attachment
                    filename = part.get_filename()
                    if filename in IGNORE_LIST:
                        # If the filename is in the ignore list, move on
                        logger.debug(f'Ignoring attachment : {filename}')
                        continue
                    # Make sure that date is actually available
                    date = None
                    try:
                        if msg.get('date',None):
                            date = time.struct_time(email.utils.parsedate_tz(msg.get('date'))[:9])
                            #date = time.strptime( msg.get('date')[:-6] ,'%a, %d  %b %Y %H:%M:%S')
                    except Exception as e:
                        logger.error(e)
                    attachment = {
                            'filename': part.get_filename(),
                            'data': part.get_payload(decode=True),
                            'from_addr': msg.get('from'),
                            'to_addr': msg.get('to'),
                            'subject': msg.get('subject'),
                            'timestamp': date
                            }
                    attachment = Attachment(attachment)
                    self.attachments.append(attachment)


class Gmail:
    def __init__(self, username, password):
        self.username = username
        self.password = password
        self._login()

    def __del__(self):
        self.logout()

    def __exit(self, msg):
        # Write error messages to syslog
        exit(1)

    def _login(self):
        try:
            logger.debug('Connecting to GMAIL..')
            self.conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
            logger.debug('Successfully connected to GMAIL')
        except imaplib.IMAP4.error as e:
            logger.error(f"Could not connect to the server : {e}")
            self.__exit(e)

        try:
            logger.debug(f'Logging in as {self.username}')
            response, data = self.conn.login(self.username, self.password)
            logger.debug('Successfully logged in ')
        except imaplib.IMAP4.error as e:
            logger.error(f"Could not authenticate : {e}")
            self.__exit(e)

        try:
            logger.debug(f'Capabilities: {self.conn.capabilities}')
            if 'UTF8' in self.conn.capabilities:
                self.conn.enable('UTF8=ACCEPT')

        except imaplib.IMAP4.error as e:
            logger.error(f'Could not read capabilities : {e}')

    def close(self):
        try:
            logger.debug('Closing mailbox')
            self.conn.close()
        except imaplib.IMAP4.error as e:
            logger.error(f'Could not close mailbox - {e}')
            self.__exit(e)

    def logout(self):
        try:
            logger.debug('Logging out of GMAIL..')
            self.conn.logout()
            logger.debug('Successfully logged out')
        except imaplib.IMAP4.error as e:
            logger.error(f'Could not logout connection : {e}')
            self.__exit(e)

    def get_mailboxes(self):
        mailboxes = self.conn.list()
        if mailboxes[0] != 'OK':
            logger.error('Could not get list of mailboxes')
            return []
        result = []
        regex = re.compile('\((?P<features>[\w\\\\ ]+)\) \"/\" \"(?P<name>[\w\[\]\/]+)\"')
        for mailbox in mailboxes[-1]:
            tmp = regex.match(mailbox)
            if not tmp:
                continue
            matchD = tmp.groupdict()
            features = []
            for feature in tmp.groupdict()['features'].split(' '):
                features.append(feature.strip('\\'))
            matchD['features'] = features
            result.append(matchD)
        return result

    def delete(self, emailid_list, expunge=False, safety=True, safety_count=10):
        """
        Delete emails based on the number
        @email_numL: str or list of email ids
        @expunge: bool; expunge mailbox after deleting the emails. Default False
        @safety: bool; Disable this to delete more than 10 emails 
        @safety_count: int; max count of emails before checking safety value
        """
        if not isinstance(emailid_list, list):
            emailid_list = [emailid_list]
        if not emailid_list:
            logger.debug('No emails to delete')
            return False
        logger.debug('Deleting %s mails' % (len(emailid_list)))
        # Safety measure to stop accidentally deleting a lot of emails 
        if len(emailid_list) > safety_count:
            logger.warn(f'About to delete more than 10 emails..')
            if safety:
                return False
        response, msg_count = self.conn.select('INBOX', readonly=False)
        if response == 'NO':
            logger.error(f'Could not select mailbox > {mailbox}')
            return False
        try:
            self.conn.store(','.join(emailid_list), '+X-GM-LABELS', '\\Trash')
            if expunge:
                self.conn.expunge()
            return True
        except Exception as e:
            logger.error(e)
        return False

    def advanced_search(self, download=False, **kwargs):
        """
        See SEARCH_KEYS for a list of avaiable search keys
        Usage can be found here: https://support.google.com/mail/answer/7190?hl=en
        """
        search_filter = ''
        for key in SEARCH_KEYS:
            value = kwargs.get(key)
            if value:
                search_filter += ' %s:%s' % (key, value)
        logger.debug('Searching for %s' % search_filter)
        return self.search(search_filter, download=download)

    def search(self, search_filter, mailbox='INBOX', readonly=True, download=False):
        """
        Search, based on advanced search operators for gmail
        https://support.google.com/mail/answer/7190?hl=en
        @readonly : Open the mailbox in read-only mode
        @download : When enabled, it will download the message body of all the emails that match 
        @returns a list of Email objects, or list of uid strings
        """
        emailL = []
        uidL = []
        assert mailbox
        assert search_filter
        logger.debug('Selecting mailbox : %s' % mailbox)
        response, msg_count = self.conn.select(mailbox, readonly=readonly)
        if response == 'NO':
            logger.error('Could not select mailbox > %s' % mailbox)
            return emailL
        logger.debug('Selection : %s, Number of Mailboxes : %s' % (response, len(msg_count)))

        result, data = self.conn.search(None, 'X-GM-RAW', '"%s"' % search_filter)
        if not data:
            logger.debug('No data returned from search')
            return emailL

        data = data[0].split()
        logger.debug('Found %d mail(s)' % (len(data)))
        if len(data) == 0:
            return emailL

        if not download:
            # we dont need to download the actual message body of the emails
            # just return the UIDs 
            nums = [x.decode('utf-8') for x in data]
            return nums

        # Download the body of the emails, and return a list of Email objects
        nums = ','.join([x.decode("utf-8") for x in data])
        response, data = self.conn.fetch(nums, '(RFC822)')
        logger.debug('Fetch response: %s' % response)

        for msg in data:
            if not isinstance(msg, tuple):
                continue
            # msg[0] = 58 (RFC822 {5990}
            num = msg[0].decode('utf-8').split(' (')[0]
            logger.debug('MessageID: %s' % num)
            uidL.append(num)
            msg = Email(msg[1].decode('utf-8'), num)
            emailL.append(msg)
        return emailL

        return uidL

    def getEmailsFromUids(self, uidL):
        for uid in uidL:
            logger.debug(f'Getting uid: {uid}')
            response, data = self.conn.fetch(uid, '(RFC822)')
            logger.debug('Fetch response: %s' % response)
            if len(data) != 2:
                logger.debug(f'Error reading the response, incorrect len')
                return
            msg = data[0]
            # msg[0] = 58 (RFC822 {5990} | 5990 is the size of the message
            num = msg[0].decode('utf-8').split(' (')[0]
            logger.debug('MessageID: %s' % num)
            yield msg

    def getAttachmentsSince(self, since, mailbox='INBOX'):
        """
        Returns a list of email objects for all messages since the specified date
        @since : strftime("%d-%b-%Y")
        @mailbox: label of the IMAP mailbox
        """
        attachments = []
        assert mailbox
        assert since
        logger.debug('Selecting mailbox : %s' % mailbox)
        response, msg_count = self.conn.select(mailbox, readonly=True)
        if response == 'NO':
            logger.error('Could not select mailbox > %s' % mailbox)
            return attachments
        logger.debug('Selection : %s, Number of Mailboxes : %s' % (response, len(msg_count)))
        result, data = self.conn.search(None, 'X-GM-RAW', '"has:attachment after:%s"' % since)
        if not data:
            return attachments

        data = data[0].split()
        logger.debug('Found %d mail(s) since %s with attachments' % (len(data), since))
        nums = ','.join([x.decode('utf-8') for x in data])

        response, data = self.conn.fetch(nums, '(RFC822)')
        for num in data:
            if not isinstance(num, tuple):
                continue
            msg = Email(num[1].decode('utf-8'), num)
            msg.extractAttachments()
            if msg.attachments:
                attachments.extend(msg.attachments)
        logger.debug('Found %s valid attachment(s) since %s' % (len(attachments), since))
        self.conn.close()
        logger.debug('Closed mailbox %s' % mailbox)
        return attachments

    def getAttachmentsForDays(self, days, mailbox='INBOX'):
        """
        Returns a list of email objects for all messages since the last X days
        @days   :   number of days to go back
        @mailbox:   label of the IMAP mailbox
        """
        # Gmail's date for using filters should be "%Y/%m/%d"
        date = (datetime.date.today() - datetime.timedelta(days)).strftime("%Y/%m/%d")
        return self.getAttachmentsSince(date, mailbox)
