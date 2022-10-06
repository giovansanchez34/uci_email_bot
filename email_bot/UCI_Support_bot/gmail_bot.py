import imapclient, smtplib, pyzmail 
from decouple import config 
from urllib.parse import unquote 
from re import finditer 
from datetime import datetime
from pytz import timezone 
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart 
import schedule



class UciInvalidEmail(Exception):
    """Raised when the given email address is not a valid uci email."""
    pass


class UciSupportBot:
    """
    A Bot that goes through the user's email inbox to look for an email sent
    from UCI Support to respond back to their Daily symptom check.
    """
    imap_server_domain = 'imap.gmail.com'
    smtp_server_domain = 'smtp.gmail.com'
    uci_domain = '@uci.edu'
    
    
    def __init__(self, email: str, password: str, response: str):
        """Initializes the state of the instance object."""
        self.email = email 
        self._password = password 
        self.response = response
        self._email_sent = False
        self._email_found = False
    
    def bot_summary(self) -> (bool, bool):
        """
        Returns a 2-tuple that contains whether or not the bot was activated 
        and if the email was sent to the recipient.
        """
        return (self._email_sent, self._email_found)
    
    @property 
    def response(self) -> str:
        """Returns the user's response."""
        return self._response 
    
    @response.setter 
    def response(self, user_response: str) -> None:
        """Ensures that user response is valid."""
        if user_response.lower() == 'not today' or user_response.lower() == 'no' or user_response.lower() == 'yes':
            self._response = user_response.lower()
        else:
            raise ValueError(f"UciSupportBot.response: given response({user_response}) must be Not Today/No/Yes")
    
    @property 
    def email(self) -> str:
        """Returns the user's email address."""
        return self._email 
    
    @email.setter 
    def email(self, email_address: str) -> None:
        """
        Stores the user's email if it ends with the domain @uci.edu, 
        has only 1 @uci.edu in the email, and the length is bigger than the
        length of @uci.edu. Otherwise it raises an UciInvalidEmail Exception.
        """
        if len(email_address) > len(self.uci_domain) and email_address.endswith(self.uci_domain)\
            and email_address.count(self.uci_domain) == 1:
                self._email = email_address
        else:
            raise UciInvalidEmail(f'UciSupportBot.email: email_address({email_address}) must contain only 1 @uci.edu,\
end with @uci.edu, and the length has to be greater than length of @uci.edu')
    
    def run_bot(self) -> None:
        """
        Runs the UCISupportBot to look for the UCI Support email at today's given
        date and sends an email to confirm our status.
        """
        imap_server, smtp_server = None, None
        try:
            imap_server = self._create_imap_server()
            smtp_server = self._create_smtp_server()
            emails = self._find_emails(imap_server)
            email_contents = self._find_email(emails)
            if email_contents is None:
                return 
            self._email_found = True 
            failed_emails = self._send_email(smtp_server, **email_contents.groupdict())
            if len(failed_emails) == 0:
                self._email_sent = True
            
        except imapclient.exceptions.LoginError:
            print("UciSupportBot.run_bot: Could not connect to a imap server. Given Credentials Invalid.")
            return
        except smtplib.socket.gaierror:
            print("UCISupportBot.run_bot: Failed to establish a connection.")
            return 
        finally:
            if imap_server is not None:
                imap_server.logout()
            if smtp_server is not None:
                smtp_server.quit()
    
    def _send_email(self, smtp_server: smtplib.SMTP, emails: str, subject: str, body: str) -> dict:
        """
        Composes an email using the given emails, subject, and body and sends it
        to the following emails through gmail.
        """
        email = MIMEMultipart()
        email['From'] = self.email 
        email['To'] = ', '.join(email for email in emails.split(','))
        email['Subject'] = unquote(subject)
        email.attach(MIMEText(unquote(body), 'plain'))
        message = email.as_string()
        return smtp_server.sendmail(self.email, emails.split(','), message)
        
        
        
        
    def _find_email(self, emails: (pyzmail.PyzMessage, )) -> str or None:
        """
        Finds an email that is sent from UCI Support, is the most recent, and contains
        a mailto link somewhere in its body. If the email meets the following criteria,
        it returns the contents of the email. If no email is found, it returns 'failed'.
        """
        for email in emails:
            parsed_links = self._parse_email(email)
            if parsed_links is not None:
                email_content = self._parse_mailto_link(parsed_links)
                if email_content is not None:
                    return email_content

                
    @staticmethod 
    def _parse_email(email: pyzmail.PyzMessage) -> 'generator of re.Match objects':
        """
        Returns None if the email is not embedded with html. Otherwise it returns
        a generator of re.match objects.
        """
        if email.html_part is not None: 
            html_text = email.html_part.get_payload().decode(email.html_part.charset)
            return finditer(r'mailto:(?P<emails>.+)\?subject=(?P<subject>.+)&amp;body=(?P<body>.+)', html_text)
        else:
            return None 
    
    def _parse_mailto_link(self, links: 'generator of re.Match objects') -> 're.Match':
        """
        Returns a re.Match object that contains the link that matches
        the user's response. If the generator links is empty, it returns None.
        """
        count = {'not today': 0, 'no': 1, 'yes': 2}.get(self.response)
        for mailto_link in links:
            if count == 0:
                return mailto_link 
            count -= 1
    
        
    @property 
    def password(self):
        """Returns the user's password."""
        return self._password 
    
    @staticmethod 
    def _find_emails(imap_server: imapclient.IMAPClient) -> (pyzmail.PyzMessage,):
        """
        Finds all emails in the user's email inbox that are from UCI Support on
        today's date and converts them into easy-to-parse PyzMessage objects.
        """
        imap_server.select_folder("INBOX", readonly=False)
        date_string = UciSupportBot._get_local_time('US/Pacific')
        email_ids = imap_server.search(['FROM','uci@service-now.com','ON', date_string])
        return tuple(pyzmail.PyzMessage.factory(email[b'BODY[]']) 
               for email in imap_server.fetch(email_ids, ['BODY[]']).values())
    
    @staticmethod 
    def _get_local_time(local_time: str) -> str:
        """
        Returns a date in DD-MM-YYYY str format in user's local time.'
        """
        today_date = datetime.now().astimezone(timezone(local_time))
        return f'{today_date.strftime("%d-%b-%Y")}'
        
        
    def _create_imap_server(self) -> imapclient.IMAPClient:
        """
        Creates the imap server needed to find the emails sent from
        UCI Support.
        """
        imap_server = imapclient.IMAPClient(self.imap_server_domain)
        imap_server.login(self.email, self.password)
        return imap_server 
    
    def _create_smtp_server(self) -> smtplib.SMTP:
        """
        Creates and establishes the smtp server needed to send a confirmation email to UCI
        Support.
        """
        smtp_server = smtplib.SMTP(self.smtp_server_domain, 587)
        smtp_server.ehlo()
        smtp_server.starttls()
        smtp_server.login(self.email, self.password)
        return smtp_server 
    

     
if __name__ == '__main__':
    EMAIL, PASSWORD = config('SUPPORT_EMAIL'), config('PASS')
    bot = UciSupportBot(EMAIL, PASSWORD, 'not today')
    bot.run_bot()
    email_sent, email_found = bot.bot_summary()
    
    if not email_sent or not email_found:
        schedule.every().hour.do(bot.run_bot)
        while not email_sent or not email_found:
            schedule.run_pending()
            email_sent, email_found = bot.bot_summary()
            
        

            
                
                
            
            
        
        
        

    
        
        
    
    
    
    

        
    
    
    
    