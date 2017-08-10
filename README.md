# Spreadsheet Mailer App

When you have the need to send slightly customized emails to a relatively medium size mailing list, you can use this script to assist you. The entire interface with the data is in google sheets. The final output is exactly equal to you sending one mail at a time. This is not a replacement for bulk mailing list management systems like mailchimp.

## Use Cases

- You have 60 guests to whom you need to send a wedding party invite with the most personal touch
- You have a list of 500 emails to invite them to a meetup
- Send weekly updates to a hand full of folks with a small customized section each week

## Key Features

- Easily tweak emails anytime in google sheet and send emails to a new list
- Add or modify templates anytime in your google sheet
- Protects against double sending emails by book keeping
- Emails you send will reflect in your "sent" mailbox as if you sent them manually

## Setup (one time)

1. Install the python libraries on your local laptop
```
pip install -r requirements.txt
```

2. Get a app password for your google account

Assume that you want to send an email out from spreadheetmailer@gmail.com. In this step
you will have to create an "app password" in google to allow access to your email account
from this script. You can create one in the following steps.

 - Enable two step verification at https://myaccount.google.com/security
 - Create a new app password as shown in the images below
 - Export two environment variables with the value (you can add it to bash profiles)

```
## mail environment variables for sending out email updates
export MAGIC_WORD='your-password-from-step b'
export FROM_GMAIL_ID='your-gmail-user-name (e.g., spreadheetmailer@gmail.com)'
```
3. Download an API credentials file for your google sheet
 - visit https://console.developers.google.com/iam-admin/projects
 - Click "Drive API" > Click "Credentials" > Click "Create Credentials" > Select "Service Account Key"
 - Create a new "Service Account" > Click "Create" (you will be assigned a temporary IAM email at this point)
 - This will download you a file. Rename it to `email-reachout-sheet-cred.json` and move it to `/etc/email-reachout-sheet-cred.json` folder
 - Give the excel sheet you care about sending emails from permissions to the temporary password user you had created. This email shoudl looks like "the-service-name-you-chose@email-reachouts.iam.gserviceaccount.com"

## Send Emails

### dry run to watch the email being sent
```
python spreadsheet-mailer-app.py --dry-run
```

### the real deal
```
python spreadsheet-mailer-app.py
```
