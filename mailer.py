import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

gmail_user = os.environ['FROM_GMAIL_ID']
gmail_password = os.environ['MAGIC_WORD']

class MailFailure(Exception):
    pass

class GmailMailer(object):
    def __init__(self):
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.ehlo()
            server.starttls()
            server.login(gmail_user, gmail_password)
            self._server = server
        except Exception as e:  
            print 'Unable to initialize gmail mailer', e

    def send_message(self, from_email, to_email, subject, message):
        # Create message container - the correct MIME type is multipart/alternative.
        msg = MIMEMultipart('alternative')

        msg['Subject'] = subject
        msg['From'] = from_email
        msg['To'] = to_email

        part = MIMEText(message, 'html')
        msg.attach(part)

        try:
            self._server.sendmail(from_email, to_email, msg.as_string())
            #print msg.as_string()
            print "Email Sent to: " + to_email
        except Exception as e:
            print "Failed to Sent Email", e
            raise MailFailure()

