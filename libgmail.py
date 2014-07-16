#!/usr/bin/python
'''
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
'''
import re
import imaplib
import email
import time
import logging
import datetime
import syslog
from hashlib import md5

# Ignore the following attachments
IGNORE_LIST = []

__version__ = '0.2.1'

class GmailException(Exception):
    pass

class Attachment(object):
    def __init__(self, data):
        self.__dict__.update(data)
        self.__calcMD5()
    
    def __calcMD5(self):
        self.md5sum = md5(self.data).hexdigest()

class Email(object):
    def __init__(self, msg_data, num, logger):
        '''
        @msg_data   : email message
        @num        : IMAP email id
        '''
        self.msg = email.message_from_string(msg_data)
        self.num = num
        self.__logger = logger
        self.attachments = []

    def extractAttachments(self):
        '''Extract attachments from an Email object
        appends to the list of Attachment objects property 'attachments'
        '''
        msg = self.msg
        for part in msg.walk():
            if part.get('Content-Disposition',None):
                if part.get('Content-Disposition').startswith('attach'):
                    # This looks like an attachment
                    filename = part.get_filename()
                    if filename in IGNORE_LIST:
                        # If the filename is in the ignore list, move on
                        self.__logger.debug('Ignoring attachment : %s'%(filename))
                        continue
                    # Make sure that date is actually available
                    date = None
                    try:
                        if msg.get('date',None):
                            date = time.struct_time(email.utils.parsedate_tz(msg.get('date'))[:9])
                            #date = time.strptime( msg.get('date')[:-6] ,'%a, %d  %b %Y %H:%M:%S')
                    except Exception,e:
                        self.__logger.error(e)
                    attachment = { 
                            'filename': part.get_filename(),
                            'data':part.get_payload(decode=True),
                            'from_addr': msg.get('from',None),
                            'to_addr':msg.get('to',None),
                            'subject': msg.get('subject',None),
                            'timestamp': date
                            }
                    attachment = Attachment(attachment)
                    self.attachments.append(attachment)

class Gmail(object):
    def __init__(self, username, password, verbose=False):
        self.username = username
        self.password = password
        self.verbose = verbose
        self._setupLogging()
        self._login()
    
    def __del__(self):
        self.logout()

    def _setupLogging(self):
        logging.basicConfig(level=logging.INFO, format="%(levelname)s %(module)s:%(funcName)s() :%(lineno)d|  %(message)s" )
        self.__logger = logging.getLogger(__name__)
        self.__logger.setLevel(logging.INFO)
        if self.verbose:
            self.__logger.setLevel(logging.DEBUG)
        # Setup syslogging

    def _error(self, msg):
        # Write error messages to syslog
        exit(1)

    def _login(self):
        try:
            self.__logger.debug('Connecting to GMAIL..')
            self.conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
            self.__logger.debug('Successfully connected to GMAIL')
        except imaplib.IMAP4.error,e:
            self.__logger.error("Could not connect to the server : %s"%(e))
            self._error(e)
        
        try:
            self.__logger.debug('Logging in as %s'%(self.username))
            response, data = self.conn.login(self.username , self.password)
            self.__logger.debug('Successfully logged in ')
        except imaplib.IMAP4.error,e:
            self.__logger.error("Could not authenticate : %s"%(e))
            self._error(e)

    def close(self):
        try:
            self.__logger.debug('Closing mailbox')
            self.conn.close()
        except imaplib.IMAP4.error,e:
            self.__logger.error('Could not close mailbox - %s'%(e))
            self._error(e)

    def logout(self):
        try:
            self.__logger.debug('Logging out of GMAIL..')
            self.conn.logout()
            self.__logger.debug('Successfully logged out')
        except imaplib.IMAP4.error,e:
            self.__logger.error('Could not logout connection : %s'%(e))
            self._error(e)
    
    def get_mailboxes(self):
        mailboxes = self.conn.list()
        if mailboxes[0] != 'OK':
            self.__logger.error('Could not get list of mailboxes')
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

    def delete(self, email_numL, expunge=True):
        '''
        Delete emails based on the number
        @email_numL: str or list of email ids
        @expunge: bool; expunge mailbox after deleting the emails
        '''
        if not isinstance(email_numL,list):
            email_numL = [email_numL]
        self.__logger.debug('Deleting %s mails'%(len(email_numL)))
        try:
            email_numL = ','.join(email_numL)
            self.conn.store(email_numL, '+FLAGS', '\Deleted')
            if expunge:
                self.conn.expunge()
        except Exception,e:
            self.__logger.error(e)
        
    def search(self, search_filter, mailbox='INBOX', has_attachment=False,
                        readonly=True):
        '''
        Search, based on advanced search operators for gmail
        https://support.google.com/mail/answer/7190?hl=en
        @readonly : Open the mailbox in read-only mode
        @returns a list of Email objects
        '''
        assert mailbox
        assert search_filter
        self.__logger.debug('Selecting mailbox : %s'%(mailbox))
        response, msg_count = self.conn.select(mailbox, readonly=readonly)
        if response == 'NO':
            self.__logger.error('Could not select mailbox > %s'%(mailbox))
            return None
        self.__logger.debug('Selection : %s, Number of Mails : %s'%(response, len(msg_count)))
        
        result, data = self.conn.search(None, 'X-GM-RAW', search_filter)
        if not data:
            self.__logger.debug('No data returned from search')
            return None

        emailL = []
        data = data[0].split()
        self.__logger.debug('Found %d mail(s)'%( len(data)))
        if len(data) == 0:
            return emailL
        nums = ','.join(data)
        
        response, data = self.conn.fetch(nums, '(RFC822)')
        self.__logger.debug('Fetch response: %s'%(response))

        for msg in data:
            if not isinstance(msg, tuple):
                continue
            # msg[0] = 58 (RFC822 {5990}
            num = msg[0].split(' (')[0]
            msg = Email(msg[1], num, self.__logger)
            emailL.append(msg)
        return emailL

    def getAttachmentsSince(self, since, mailbox='INBOX'):
        '''
        returns:
            list of attachments if successful
            None if error/Failure
        '''
        try:
            return self._getAttachmentsSince(since, mailbox)
        except Exception,e:
            self.__logger.error(e)
        return None

    def _getAttachmentsSince(self, since, mailbox='INBOX'):
        '''
        Returns a list of email objects for all messages since the specified date
        @since : strftime("%d-%b-%Y")
        @mailbox: label of the IMAP mailbox
        '''
        assert mailbox
        assert since
        self.__logger.debug('Selecting mailbox : %s'%(mailbox))
        response, msg_count = self.conn.select(mailbox, readonly=True)
        if response == 'NO':
            self.__logger.error('Could not select mailbox > %s'%(mailbox))
            return None
        self.__logger.debug('Selection : %s, Number of Mailboxes : %s'%(response, len(msg_count)))
        #result, data = self.conn.search(None, '(SINCE %s)'%(since))
        result, data = self.conn.search(None, 'X-GM-RAW', 'has:attachment after:%s' %(since))
        if not data:
            return None

        data = data[0].split()
        self.__logger.debug('Found %d mail(s) since %s with attachments'%( len(data), since))
        nums = ','.join(data)

        attachments = []
        response, data = self.conn.fetch(nums, '(RFC822)')
        for num in data:
            if not isinstance(num, tuple):
                continue
            msg = Email(num[1], self.__logger)
            msg.extractAttachments()
            attachments.append(msg.attachments)
        self.__logger.debug('Found %s valid attachment(s) since %s'%( len(attachments), since ))
        self.conn.close()
        self.__logger.debug('Closed mailbox %s'%(mailbox))
        return attachments

    def getAttachmentsForDays(self, days, mailbox='INBOX'):
        '''
        Returns a list of email objects for all messages since the last X days
        @days   :   number of days to go back
        @mailbox:   label of the IMAP mailbox
        '''
        # The date should be in the IMAP RFC Format - "%d-%b-%Y"
        #date = (datetime.date.today() - datetime.timedelta(days)).strftime("%d-%b-%Y")
        # Gmail's date for using filters should be "%Y/%m/%d"
        date = (datetime.date.today() - datetime.timedelta(days)).strftime("%Y/%m/%d")
        return self.getAttachmentsSince( date, mailbox )
