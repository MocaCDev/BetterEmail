import imaplib
import email
from email.parser import HeaderParser
import html2text
import sys
from colorama import Fore, Style
import json

class BetterEmail:

    IMAP_GMAIL = ('imap.gmail.com', 993)
    
    # Months that will be set to `self.since`
    JAN = '1-Jan-{}'
    FEB = '1-Feb-{}'
    MAR = '1-Mar-{}'
    APR = '1-Apr-{}'
    MAY = '1-May-{}'
    JUN = '1-Jun-{}'
    JUL = '1-Jul-{}'
    AUG = '1-Aug-{}'
    SEP = '1-Sep-{}'
    OCT = '1-Oct-{}'
    NOV = '1-Nov-{}'
    DEC = '1-Dec-{}'

    def __init__(self, email, password, imap: tuple):
        self.email_connection = imaplib.IMAP4_SSL(imap[0], imap[1])

        try:
            self.user = self.email_connection.login(email, password)
        except imaplib.IMAP4.error as IE:
            sys.stderr.write(f'{Fore.RED}BetterEmail Init Error:{Style.RESET_ALL}\n\tFailed to login to {email}.\n\n')
            sys.stderr.write(f'Full Error:\n\t{Fore.RED}{str(IE)}{Style.RESET_ALL}\n\n')
            sys.stderr.flush()
            sys.exit(1)
        
        self.mailparser = HeaderParser()

        # Inbox is selected by default
        self.email_connection.select('"Inbox"')

        # Application-specific variables
        self.since = ''
        self.category = '"Inbox"' # defaults to inbox

    def set_since(self, since): self.since = f'SINCE {since} X-GM-RAW'
    def set_category(self, category): self.category = f'"category:{category}"'

    def grab_emails(self, max=20):

        # Might be removed for actions can still occurr even without a date specified
        if self.since == '':
            sys.stderr.write(f'{Fore.RED}BetterEmail Error:{Style.RESET_ALL}\n\tCannot grab messages from a span of infinite time.\n\tSet the `since` date by using {Fore.GREEN}`BetterEmail.set_since`{Style.RESET_ALL}.\n\tExample: {Fore.CYAN}BetterEmail.set_since("1-Dec-2022"){Style.RESET_ALL}\n\n')
            sys.stderr.flush()
            sys.exit(1)

        result, messages = self.email_connection.uid('search', f'{self.since} {self.category}')

        #messages_returned = [messages[0].split()[i] for i in range(len(messages[0].split())) if i <= max]

        return result, messages#messages_returned
    
    def get_body(self, messages = [], max=20):
        if len(messages) == 0:
            return
        
        email_body = {
            'From': '',
            'To': '',
            'Subject': '',
            'Date': '',
            'Body': ''
        }

        all_ = []
        pre = []
        
        for i in messages:
            if len(all_) >= max:
                break

            try:
                result, email_data = self.email_connection.uid('fetch', i, '(RFC822)')
            except:
                sys.stderr.write(f'{Fore.RED}BetterEmail Error:{Style.RESET_ALL}\n\tCould not fetch data.\n')
                sys.stderr.flush()
                sys.exit(1)
            
            if result == 'OK':
                for response in email_data:
                    if isinstance(response, tuple):
                        email_message = email.message_from_bytes(response[1])

                        email_body['From'] = email_message['From']
                        email_body['To'] = email_message['To']
                        email_body['Subject'] = email_message['Subject']
                        email_body['Date'] = email_message['Date']
                
                m = email.message_from_string(email_data[0][1].decode('utf-8'))

                for p in m.walk():
                    if p.get_content_type() == 'text/plain':
                        text = p.get_payload(decode=True)
                        ct = p.get_content_type()
                        cc = p.get_content_charset() # charset in Content-Type header
                        cte = p.get("Content-Transfer-Encoding")
                        #print("part: " + str(ct) + " " + str(cc) + " : " + str(cte))

                        try:
                            text = text.decode('utf-8')
                            if '&zwnj;' in text:
                                text = text.replace('&zwnj;', '')
                        except:
                            try:
                                text = text.decode('windows-1254')
                                text = text.encode('ascii', 'ignore').decode('unicode_escape')
                            except:
                                pass

                        text = text.split('\r')
                        cont = True
                        for i in range(len(text)):
                            if not cont: break

                            if text[i].replace(' ', '') == '\r' or text[i].replace(' ', '') == '\n' or text[i].replace(' ', '') == '':
                                del text[0]
                            else:
                                cont = False
                        
                        text = ''.join(text)
                        email_body['Body'] = text
                        
                all_.append(email_body)

        return all_

    
    def grab_emails_body(self, max=20):
        r, m = self.grab_emails(max)
        email_bodies = self.get_body(m, max)

        for i in email_bodies:
            print(f'From: {i["From"]}')
            print(f'To: {i["To"]}')
            print(f'Subject: {i["Subject"]}')
            print(f'Date: {i["Date"]}')
            print(f'Body:\n{i["Body"]}', end='\n\n')
        
    def grab_emails_body_and_dump(self, dump_file='dump.json', max=20):
        r, m = self.grab_emails(max)
        email_bodies = self.get_body(m, max)

        with open(dump_file, 'w') as file:
            file.write(json.dumps(email_bodies, indent=2))
            file.flush()
            file.close()