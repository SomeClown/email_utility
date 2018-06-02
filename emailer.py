#!/usr/bin/env python3

"""
Emailer CLI
"""

import click
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, GSSAPI, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version, FolderCollection

__author__ = "SomeClown"
__license__ = "MIT"
__maintainer__ = "Teren Bryson"
__email__ = "teren@wwt.com"

EPILOG = 'Don\'t do anything I wouldn\'t do'
CONTEXT_SETTINGS = dict(help_option_names=['-h', '--help'])


@click.group(epilog=EPILOG, context_settings=CONTEXT_SETTINGS)
def cli():
    """
    Quick and dirty email utility to do my bidding

    """
    pass


def load_email_from_file(filename: str):
    """
    Load email content from a file
    :param filename: 
    :return: 
    """
    with open(filename) as f:
        content = f.read()
    return content

@click.command(short_help='Email wizardry')
@click.option('-f', '--from', 'from_addr', help='from address')
@click.option('-t', '--to', 'to_addr', help='to address')
@click.option('-s', '--subject', 'sub', help='subject')
@click.option('-b', '--body', 'body', help='body content')
@click.option('-a', '--act', 'account', help='account')
@click.option('-u', '--user', 'user_name', help='account username')
@click.option('-p', '--pwd', 'pwd', help='pwd')
@click.option('--sp', 'smtp_port', help='outgoing port')
def do_normal_email(from_addr: str, to_addr: str, sub: str, body: str, account: str, pwd: str, user_name: str, smtp_port=587):
    """
    
    Sends an email (this is mostly for testing. The Exchange stuff below is what I'm mostly developing)
    
    :param from_addr: 
    :param to_addr: 
    :param sub: 
    :param body: 
    :param account:
    :param pwd:
    :param user_name:
    :param smtp_port:
    :return: 
    """
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = sub

    msg.attach(MIMEText(body, 'plain'))

    server = smtplib.SMTP(account, smtp_port)
    server.starttls()
    server.login(user_name, pwd)
    text = msg.as_string()
    server.sendmail(from_addr, to_addr, text)
    server.quit()


@click.command(short_help='Email wizardry')
@click.option('-f', '--from', 'from_addr', help='from address')
@click.option('-a', '--act', 'account', help='account')
@click.option('-u', '--user', 'user_name', help='account username')
@click.option('-p', '--pwd', 'pwd', help='pwd')
def get_exchange_email(from_addr: str, account: str, pwd: str, user_name: str):
    """

    Gets email with Exchange

    :param from_addr: 
    :param account:
    :param pwd:
    :param user_name:
    :return: 
    """
    creds = Credentials(user_name, pwd)
    config = Configuration(server=account, credentials=creds)
    account = Account(primary_smtp_address=from_addr,
                      autodiscover=False, access_type=DELEGATE, config=config)
    for item in account.inbox.all().order_by('-datetime_received')[:10]:
        print(item.subject, item.sender, item.datetime_received)


@click.command(short_help='Email wizardry')
@click.option('-f', '--from', 'from_addr', help='from address')
@click.option('-t', '--to', 'to_addr', help='to address')
@click.option('-s', '--subject', 'sub', help='subject')
@click.option('-b', '--body', 'body', help='body content')
@click.option('-a', '--act', 'account', help='account')
@click.option('-u', '--user', 'user_name', help='account username')
@click.option('-p', '--pwd', 'pwd', help='pwd')
@click.option('--file', 'file', help='file')
def send_exchange_email(from_addr: str, to_addr: str, sub: str, body: str,
                        account: str, pwd: str, user_name: str, file: str):
    """

    Sends email using Exchange

    :param from_addr: 
    :param to_addr: 
    :param sub: 
    :param body: 
    :param account:
    :param pwd:
    :param user_name:
    :param file:
    :return: 
    """
    if file:
        body = load_email_from_file(file)
        print(body)
    else:
        pass
    creds = Credentials(user_name, pwd)
    config = Configuration(server=account, credentials=creds)
    account = Account(primary_smtp_address=from_addr,
                      autodiscover=False, access_type=DELEGATE, config=config)
    msg = Message(
        account=account,
        folder=account.sent,
        subject=sub,
        body=body,
        to_recipients=[Mailbox(email_address=to_addr)])
    msg.send_and_save()

cli.add_command(do_normal_email, 'email')
cli.add_command(get_exchange_email, 'get_exchange')
cli.add_command(send_exchange_email, 'send_exchange')

if __name__ == '__main__':
    cli()
