#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import argparse
import logging
import logging.handlers
from O365 import Account
from O365 import MSGraphProtocol
from O365 import FileSystemTokenBackend
import config

"""
    This script uses the O365 library to connect
    to Office 365 with the MSGraphProtocol and the
    modern authentification standard OAuth2.0-Bearer
"""

"""
    ToDo's
    - Change Logging to more convenient version
    - implement single file push
    - improve mail sending
"""


# Here comes your (few) global variables
PROG = os.path.basename(sys.argv[0])


# setup logging
logger = logging.getLogger(PROG)
logger.setLevel(logging.DEBUG)  # global logger, no restrictions

# create file handler which logs even debug messages
fh = logging.FileHandler(config.LOG_PATH, mode='a')
fh.setLevel(logging.DEBUG)
fh.set_name('File')

# create formatter and add it to the handlers
formatter = logging.Formatter(
  "%(asctime)s %(name)-20s - %(funcName)-20s - %(levelname)-8s  - %(message)s",
  datefmt='%y-%m-%d %H:%M:%S')
fh.setFormatter(formatter)

# add the handlers to the logger
logger.addHandler(fh)


def ensure_absolute_path(my_path: str):
    """Make absolut path based on executing directory"""
    if not os.path.isabs(my_path):
        cwd = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(cwd, my_path)
    else:
        return my_path


def make_folder(folder, mod=0o600):
    """Create Folder if it does not exist"""
    absFolder = ensure_absolute_path(folder)
    if not os.path.exists(absFolder):
        os.mkdir(absFolder, mod)
        logger.debug("Folder %s created", absFolder)
    return absFolder


def safe_file_name(filename, replace=' '):
    """Make safe filename"""
    import unicodedata
    import string

    valid_filename_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    char_limit = 150  # 255 replaced by 150 to be onsafe side

    # replace spaces
    for r in replace:
        filename = filename.replace(r, '_')

    # keep only valid ascii chars
    cleaned_filename = unicodedata.normalize(
        'NFKD',
        filename).encode('ASCII', 'ignore').decode()

    # keep only whitelisted chars
    cleaned_filename = ''.join(
        c for c in cleaned_filename if c in valid_filename_chars)
    if len(cleaned_filename) > char_limit:
        logger.warning("Warning, filename truncated because it was over {}. Filenames may no longer be unique".format(char_limit))
    return cleaned_filename[:char_limit]


def parse_arguments(args):
    """Parse/define command line arguments."""
    parser = argparse.ArgumentParser(description=f'{__doc__}', formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version='0.1.0', help='Print script version.')
    parser.add_argument('-a', '--auth', action='store_true', default=False, help='Get initial or refresh token if authentification expired.')
    parser.add_argument('-k', '--keep', action='store_true', default=False, help='Keep messages after pushing to MDA.')
    parser.add_argument('-v', '--verbose', action='store_true', default=False, help='Output logger infromation to Screen.')
    parser.add_argument('-m', '--message', default=None, help='Email message as ''*.eml'' to push to RT.')

    return parser.parse_args(args)


def check_for_folders():
    """Create folder if ist does not exist."""
    make_folder(config.MAIL_PATH, 0o644)
    make_folder(config.TOKEN_PATH, 0o600)
    make_folder(os.path.dirname(os.path.abspath(config.LOG_PATH)), 0o644)


def getAccount(user_id):
    """Get account by user"""
    # prepare token backend for user
    token_backend = FileSystemTokenBackend(token_path=config.TOKEN_PATH, token_filename=user_id + '.token')

    # prepare MSGraphProtocol for user
    my_protocol = MSGraphProtocol(config.API_VERSION, user_id)

    # setup account definition for user
    return Account(credentials=(config.CLIENT_ID, config.CLIENT_SECRET), protocol=my_protocol, scopes=config.SCOPES, token_backend=token_backend)


def reauth_token(opt):
    """Initial or refresh token"""
    for n in range(0, len(config.USERS)):
        user = config.USERS[n]
        logger.debug("Requesting token for %s", user['user_id'])
        try:
            # create account
            account = getAccount(user['user_id'])
            if not account.is_authenticated:
                account.authenticate()
                logger.debug("Token for %s (%s) has been created.", user['user_id'], account.con.token_backend.token_path)
                if opt.verbose:
                    logger.info("Token for %s has been created.", user['user_id'])
            else:
                account.connection.refresh_token()
                logger.debug("Token for %s (%s) has been refreshed.", user['user_id'], account.con.token_backend.token_path)
                if opt.verbose:
                    logger.info("Token for %s has been refreshed.", user['user_id'])
        except Exception as ex:
            logger.exception('Prozedure reauth_token threw an error.\n{}'.format(ex))


def get_messages_cnt(inbox, user_id, verbose):
    """Get messages count."""
    total_items = inbox.get_messages(limit=9999)
    total_items_count = sum(1 for m in total_items)
    unread_items = inbox.get_messages(limit=9999, query='isRead eq false')
    unread_items_count = sum(1 for m in unread_items)

    if verbose:
        logger.info('{}: Seen {} messages. {} messages are unread.'.format(user_id, total_items_count, unread_items_count))
    logger.debug('{}: Seen {} messages. {} messages are unread.'.format(user_id, total_items_count, unread_items_count))

    return total_items_count, unread_items_count


def get_messages(opt):
    """Pull messages from o365"""
    for n in range(0, len(config.USERS)):
        user = config.USERS[n]
        mail_folder = os.path.join(config.MAIL_PATH, user['user_id'])
        make_folder(mail_folder, 0o644)
        logger.debug("Message pull initialized for user_id: %s", user['user_id'])
        try:
            account = getAccount(user['user_id'])
            if not account.is_authenticated:
                logger.debug("The account %s needs a reauthentication", user['user_id'])
            else:
                mailbox = account.mailbox()
                inbox = mailbox.inbox_folder()

                total, unread = get_messages_cnt(inbox, user['user_id'], opt.verbose)

                msg_cnt = 0
                # for each unread message do (25 at a time by default)
                for message in inbox.get_messages(query='isRead eq false', download_attachments=True):
                    msg_cnt += 1
                    if opt.verbose:
                        logger.info('{}: Working on message from:<{}> subject:{}.'.format(msg_cnt, message.sender, message.subject))

                    # email creation date
                    created = message.created.strftime("%Y%m%d_%H%M%S")

                    # create unique file absolut path and name
                    safe_filename = safe_file_name('{}_{}_{}'.format(created, message.sender.address, message.subject))
                    msg_abs_path = os.path.join(mail_folder, '{}.eml'.format(safe_filename))

                    # store file
                    try:
                        ret = message.save_as_eml(to_path=msg_abs_path)
                        if not ret:
                            logger.debug("From:<{}>\nSubject:{}\nCreated Date:{}".format(message.sender, message.subject, created))
                    except FileNotFoundError:
                        try:  # try rename
                            msg_abs_path = os.path.join(mail_folder, '{}_{}.eml'.format(created, message.conversation_id))
                            ret = message.save_as_eml(to_path=msg_abs_path)
                        except FileNotFoundError:
                            logger.debug("FileNotFoundError")
                            logger.debug('From:<{}>\nSubject:{}\nCreated Date:{}'.format(message.sender, message.subject, created))
                    except Exception as ex:
                        logger.debug("MessageSaveError")
                        logger.debug('From:<{}>\nSubject:{}\nCreated Date:{}'.format(message.sender, message.subject, created))

                    message.mark_as_read()

        except Exception as ex:
            logger.exception('Procedure get_messages threw an error.\n{}'.format(ex))


def get_files_in_folder(folder):
    """Return files from folder"""
    return [fn for fn in os.listdir(folder) if fn.lower().endswith('.eml')]


def push_specific_message(abs_filename, verbose=False, keep=False):
    """Push from command line"""
    import email
    import re

    f = open(abs_filename, "rb")
    message = email.message_from_binary_file(f)
    f.close()
    headers = message._headers

    to_addr = ''

    for h in headers:
        if h[0] == 'To':
            # multiple addresses possible
            to_addr = (re.findall(r'[\w\.-]+@[\w\.-]+', h[1]))

    for u in config.USERS:
        if u['user_id'] in to_addr:
            push_message(abs_filename, u, verbose, keep)


def push_message(abs_filename, user, verbose=False, keep=False):
    """Push messages"""
    import subprocess

    logger.debug("Pushing: %s", abs_filename)
    if verbose:
        logger.info("Pushing: %s", abs_filename)

    try:
        p1 = subprocess.Popen(['cat', abs_filename], stdout=subprocess.PIPE)
        p2 = subprocess.Popen(['/opt/rt4/bin/rt-mailgate --queue ''{}'' --action {} --url ''{}'' --ca-file ''{}'''.format(
            user['queue'], user['action'], config.RT_URL, config.CA_FILE)],
            stdin=p1.stdout, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        p1.stdout.close()
        output = p2.communicate()

        if output[1] != b'':
            logger.error("Error pushing '{}' onwards.".format(abs_filename))
            os.rename(abs_filename, '{}.error'.format(abs_filename))
        else:
            logger.debug("Pushed '{}' onwards.".format(abs_filename))

            if keep:
                os.rename(abs_filename, '{}.keep'.format(abs_filename))
            else:
                os.remove(abs_filename)
    except Exception as ex:
        logger.exception('Procedure push_message threw an error.\n%s', ex)


def push_messages(opt):
    """Push messages onwards"""
    for n in range(0, len(config.USERS)):
        user = config.USERS[n]
        mail_folder = os.path.join(config.MAIL_PATH, user['user_id'])
        logger.debug("Message push for user_id: %s and folder: %s",
                     user['user_id'], mail_folder)
        if opt.verbose:
            logger.info("Message push for user_id: %s", user['user_id'])

        if not os.path.exists(mail_folder):
            logger.debug("Folder '{}' does not exist. No Mails in Inbox? New Account?".format(mail_folder))
            continue

        files = get_files_in_folder(mail_folder)
        logger.debug("%s Message found to push.", len(files))

        for f in files:
            abs_filename = os.path.join(mail_folder, f)
            push_message(abs_filename, user, opt.verbose, opt.keep)


def main(args) -> None:
    """Main prozedure."""
    logger.debug("Entered main procedure.")
    logger.debug("Try parsing arguments.")
    opt = parse_arguments(args)
    logger.debug('\t\toptions: %s', opt)

    if opt.verbose:
        # create console handler with a higher log level
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.set_name('Console')

        # create formatter and add it to the handlers
        formatter = logging.Formatter("%(name)-20s: %(levelname)-8s %(message)s")
        ch.setFormatter(formatter)

        # add the handlers to the logger
        # logging.getLogger('').addHandler(ch)
        logger.addHandler(ch)

    if opt.message:
        push_specific_message(opt.message, opt.verbose, opt.keep)
        sys.exit(0)

    # Check if required folders exist
    check_for_folders()

    # forced token request and refresh
    if opt.auth:
        reauth_token(opt)

    # get messages from o365
    get_messages(opt)

    # push messages
    push_messages(opt)


if __name__ == '__main__':
    """Entrypoint."""
    try:
        logger.debug('Executing script: %s', PROG)
        main(sys.argv[1:])
    except Exception as ex:
        logger.exception('{} exception during startup: {}', PROG, ex)
        sys.stderr.write(f'{PROG}: {ex}')
        sys.exit(1)
    sys.exit(0)
