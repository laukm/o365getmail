"""
    Configuration settings for running the o365getmail script.

    In productive system, ensure file can't be accessed by unpriviledged.
"""

import os

CWD = os.path.dirname(os.path.abspath(__file__))


CLIENT_ID = ''
CLIENT_SECRET = ''


MAIL_PATH = os.path.join(CWD, 'mails')
TOKEN_PATH = os.path.join(CWD, 'tokens')
LOG_PATH = os.path.join(CWD, 'o365_email_getter.log')



# AUTHORITY_URL ending determines type of account that can be authenticated:
# /organizations = organizational accounts only
# /consumers = MSAs only (Microsoft Accounts - Live.com, Hotmail.com, etc.)
# /common = allow both types of accounts
AUTHORITY_URL = 'https://login.microsoftonline.com/common'

AUTH_ENDPOINT = '/oauth2/v2.0/authorize'
TOKEN_ENDPOINT = '/oauth2/v2.0/token'

RESOURCE = 'https://graph.microsoft.com/'
API_VERSION = 'beta'
#['basic', 'message_all']
SCOPES = ['User.Read', 'offline_access', 'Mail.ReadWrite', 'Mail.Send'] # Add other scopes/permissions as needed.


# Getter definitions for message pull
USERS = []
USERS.append({"user_id":"EMAIL@OUTLOOK.COM", "queue":"Microsurgery", "action":"correspond"})
#USERS.append({"user_id":"EMAIL1@OUTLOOK.COM", "queue":"Ophthalmic", "action":"correspond"})


# MDA settings
RT_URL = ''
CA_FILE = '/usr/local/share/ca-certificates/yourCertificate.cer'
# "MDA": "/opt/rt4/bin/rt-mailgate --queue 'Microsurgery' --action correspond --url https://dev-med-rt.zeiss.com/ --ca-file /usr/local/share/ca-certificates/dev-med-rt.zeiss.com/dev_med_rt.zeiss.com.cer"



# This code can be removed after configuring CLIENT_ID and CLIENT_SECRET above.
if 'ENTER_YOUR' in CLIENT_ID or 'ENTER_YOUR' in CLIENT_SECRET or 'ENTER_YOUR' in MAIL_PATH or 'ENTER_YOUR' in TOKEN_PATH or 'ENTER_YOUR' in LOG_PATH:
    print('ERROR: config.py does not contain valid CLIENT_ID and CLIENT_SECRET')
    import sys
    sys.exit(1)
