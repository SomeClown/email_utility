#!/usr/bin/env python3

"""

Email CLI

This mostly exists because I needed to send bulk emails to people and I didn't like the existing
solutions available, and because many of the ones that would work aren't free. This is a simple
solution for my needs. Any extra functionality beyond bulk mailing is only there because I got bored
and figured why the hell not.

Note that many of the imports below are speculative; they're there so I don't forget them
later as I add more functionality.

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


def content_from_file(filename: str):
    """
    Load email content from a file
    :param filename: 
    :return: 
    """
    with open(filename) as f:
        content = f.read()
    return content


def spreadsheet_data(spreadsheet_name, sheet_name):
    """
    Get names and other data from an Excel file
    :param spreadsheet_name:
    :param sheet_name:
    :return: 
    """
    import pandas as pd
    df = pd.read_excel(spreadsheet_name, sheet_name)
    people_list = []
    for i in df.index:
        people = {'First_Name': df['First_Name'][i], 'Last_Name': df['Last_Name'][i],
                  'Middle_Initial': df['Middle_Initial'][i], 'Email_Address': df['Email_Address'][i],
                  'Phone': df['Phone'][i], 'City': df['City'][i], 'State': df['State'][i]}
        people_list.append(people)
    return people_list


@click.command(short_help='Email wizardry')
@click.option('-f', '--from', 'from_addr', help='from address')
@click.option('-t', '--to', 'to_addr', help='to address')
@click.option('-s', '--subject', 'sub', help='subject')
@click.option('-b', '--body', 'body', help='body content')
@click.option('-a', '--act', 'account', help='account')
@click.option('-u', '--user', 'user_name', help='account username')
@click.option('-p', '--pwd', 'pwd', help='pwd')
@click.option('--sp', 'smtp_port', help='outgoing port')
def do_normal_email(from_addr: str, to_addr: str, sub: str, body: str,
                    account: str, pwd: str, user_name: str, smtp_port=587):
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
@click.option('-n', '--num', 'num_emails', help='number of emails')
def get_exchange_email(from_addr: str, account: str, pwd: str, user_name: str, num_emails: int):
    """

    Gets email with Exchange

    :param from_addr: 
    :param account:
    :param pwd:
    :param user_name:
    :param num_emails:
    :return: 
    """
    if not num_emails:
        num_emails = 10
    creds = Credentials(user_name, pwd)
    config = Configuration(server=account, credentials=creds)
    account = Account(primary_smtp_address=from_addr,
                      autodiscover=False, access_type=DELEGATE, config=config)
    for item in account.inbox.all().order_by('-datetime_received')[:int(num_emails)]:
        print(item.subject, item.sender, item.datetime_received)


@click.command(short_help='Email wizardry')
@click.option('-f', '--from', 'from_addr', help='from address')
@click.option('-t', '--to', 'to_addr', help='to address')
@click.option('-s', '--subject', 'sub', help='subject')
@click.option('-b', '--body', 'body', help='body content')
@click.option('-a', '--act', 'account', help='account')
@click.option('-u', '--user', 'user_name', help='account username')
@click.option('-p', '--pwd', 'pwd', help='pwd')
def send_exchange_email(from_addr: str, to_addr: str, sub: str, body: str,
                        account: str, pwd: str, user_name: str):
    """

    Sends email using Exchange

    :param from_addr: 
    :param to_addr: 
    :param sub: 
    :param body: 
    :param account:
    :param pwd:
    :param user_name:
    :return: 
    """
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


@click.command(short_help='Bulk Email wizardry')
@click.option('-f', '--from', 'from_addr', help='from address')
@click.option('-s', '--subject', 'sub', help='subject')
@click.option('-a', '--act', 'account', help='account')
@click.option('-u', '--user', 'user_name', help='account username')
@click.option('-p', '--pwd', 'pwd', help='pwd')
@click.option('-c', '--content', 'content_file', help='file')
@click.option('--names', 'names_list', help='names list')
@click.option('--sheet', 'sheet', help='Excel Sheet')
@click.option('--signature', 'signature', help='Closing signature')
def bulk_send_exchange_email(from_addr: str, sub: str, account: str,
                             pwd: str, user_name: str, names_list: str,
                             content_file: str, sheet: str, signature: str):
    """

    Sends email using Exchange

    :param from_addr: 
    :param sub: 
    :param account:
    :param pwd:
    :param user_name:
    :param content_file:
    :param names_list:
    :param sheet:
    :param signature:
    :return: 
    """
    recipients = spreadsheet_data(names_list, sheet)
    body = content_from_file(content_file)
    signature = content_from_file(signature)
    creds = Credentials(user_name, pwd)
    config = Configuration(server=account, credentials=creds)
    account = Account(primary_smtp_address=from_addr,
                      autodiscover=False, access_type=DELEGATE, config=config)
    for item in recipients:
        complete_email = 'Hello ' + item['First_Name'] + '\n\n' + body + '\n' + signature
        email_address = item['Email_Address']
        msg = Message(
            account=account,
            folder=account.sent,
            subject=sub,
            body=complete_email,
            to_recipients=[Mailbox(email_address=email_address)])
        #print(repr(msg))
        msg.send_and_save()

cli.add_command(do_normal_email, 'email')
cli.add_command(get_exchange_email, 'get_exchange')
cli.add_command(send_exchange_email, 'send_exchange')
cli.add_command(bulk_send_exchange_email, 'bulk_send')

if __name__ == '__main__':
    cli()
