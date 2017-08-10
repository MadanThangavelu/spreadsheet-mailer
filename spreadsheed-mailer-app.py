import gspread
import datetime
import time
from jinja2 import Template
from oauth2client.service_account import ServiceAccountCredentials
from mailer import GmailMailer, MailFailure

scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]

credentials_file = '/etc/email-reachout-sheet-cred.json'

class SpreadsheetMailerApp(object):
    def __init__(self, options):
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
                            options["credentials_file"],
                            options["scope"]
                       )
        self.gc = gspread.authorize(credentials)
        # Sheets where the data is retrieved
        self.gsheet_key = options["gsheet_key"] or "15Qc7vHK91k70qDTWKTrOlnc85msNKutUuUBcwzO-64U"
        self._google_sheet = self.gc.open_by_key(self.gsheet_key)

        # Hard coded values for sheet column numbers and sheet names
        self._variable_title_row = 2
        self.email_template_sheet = options["email_template_sheet"] or "emailTemplate1"
        self.mailing_list_data_sheet = options["mailing_list_data_sheet"] or "MailingList"

        # Gather basic data from spreadsheet
        self._bookeeping_column_map = None
        self._template_variables = None
        self._mail_template = None
        self._from_email = None
        self._email_subject = None

        self.get_bookeeping_columns()
        self.get_template_variable_names()
        self.get_mail_template()
        self.get_from_email()
        self.get_subject()

        self._sending = 'SENDING'
        self._sent = 'SENT'
        self._failed = 'FAILED'

        self.mail_id_column = self._bookeeping_column_map["email"]
        self.mail_sent_date_column = self._bookeeping_column_map["mail-sent-date"]
        self.mail_status_column = self._bookeeping_column_map["mail-status"]

        self._gmail_mailer = GmailMailer()

    def get_bookeeping_columns(self):
        bookeeping_columns = ["mail-sent-date", "mail-status", "email"]
        worksheet = self._google_sheet.worksheet(self.mailing_list_data_sheet)
        column_names = worksheet.row_values(self._variable_title_row)

        bookeeping_column_map = {}

        for idx, column_name in enumerate(column_names):
            if column_name in bookeeping_columns:
                bookeeping_column_map[column_name] = idx + 1

        if not self._bookeeping_column_map:
            self._bookeeping_column_map = bookeeping_column_map

        return self._bookeeping_column_map

    def get_email_list(self):
        worksheet = self._google_sheet.worksheet(self.mailing_list_data_sheet)
        list_of_emails = worksheet.col_values(self.mail_id_column)

        # Remove the first two rows since they are custom headers
        return [i for i in list_of_emails[2:] if i != ""]

    def get_subject(self):
        worksheet = self._google_sheet.worksheet(self.email_template_sheet)
        email_subject = worksheet.acell('B3').value

        # Cache to not make call to google sheet
        if not self._email_subject:
            self._email_subject = email_subject

        return self._email_subject

    def get_from_email(self):
        worksheet = self._google_sheet.worksheet(self.email_template_sheet)
        from_email = worksheet.acell('B2').value

        # Cache to not make call to google sheet
        if not self._from_email:
            self._from_email = from_email

        return self._from_email

    def get_mail_template(self):
        worksheet = self._google_sheet.worksheet(self.email_template_sheet)
        mail_template = worksheet.acell('B1').value

        # Cache to not make call to google sheet
        if not self._mail_template:
            self._mail_template = Template(mail_template)

        return self._mail_template

    def get_template_variable_names(self):
        worksheet = self._google_sheet.worksheet(self.mailing_list_data_sheet)
        list_of_variables = worksheet.row_values(self._variable_title_row)
        template_variables = [(val[9:], idx) for idx, val in enumerate(list_of_variables) if "variable" in val]

        # Cache to not make call to google sheet
        if not self._template_variables:
            self._template_variables = template_variables

        return self._template_variables
    
    def render_single_email(self, variables):
        rendered_data = self._mail_template.render(**variables)
        return rendered_data

    def send_out_emails(self):
        worksheet = self._google_sheet.worksheet(self.mailing_list_data_sheet)
        row_columns = worksheet.get_all_values()
        for idx, row in enumerate(row_columns[2:]):
            # get email address
            to_email = row[self.mail_id_column - 1]
            if to_email == "":
                ## Account for empty rows
                print "no email found, pass"
                continue

            if row[self.mail_status_column-1] == self._sent:
                print "Skipping sending to: " + to_email + " as mail is already sent"
                continue

            variables = {}
            for variable_name, column_no in self._template_variables:
                variables[variable_name] = row[column_no]

            email_body = self.render_single_email(variables)
            try:
                self._gmail_mailer.send_message(self._from_email, to_email, self._email_subject, email_body)

                ts = time.time()
                current_date_time = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')
                worksheet.update_cell(idx + 1 + 2, self.mail_sent_date_column, current_date_time)
                worksheet.update_cell(idx + 1 + 2, self.mail_status_column, self._sent)
            except MailFailure as e:
                worksheet.update_cell(idx + 2, self.mail_status_column,  self._failed)
                print "Email sending Failed", e
            except Exception as e:
                print "Unknown error", e

if __name__ == "__main__":
    ama = SpreadsheetMailerApp({
        "mail_id_column": 1,
        "custom_message_column": 2,
        "email_sent_date": "M",
        "mailing_list_data_sheet": "MailingList",
        "email_template_sheet": "emailTemplate1",
        "gsheet_key": "15Qc7vHK91k70qDTWKTrOlnc85msNKutUuUBcwzO-64U",
        "credentials_file": credentials_file,
        "scope": scope
    })

    ama.send_out_emails()

