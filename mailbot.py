import os
import re
import base64
import httplib2
import mimetypes

from apiclient import discovery
from apiclient import errors
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


# If modifying these scopes, delete your previously saved credentials at ~/.credentials/
SCOPES = 'https://www.googleapis.com/auth/gmail.modify'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'CMPT412 Downloader'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'downloader.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def get_list_of_messages(service, user_id='me', query=''):
    """List all Messages of the user's mailbox matching the query.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me" can be used to indicate the authenticated user.
        query: String used to filter messages returned.

    Returns:
        List of Messages that match the criteria of the query. Note that the
        returned list contains Message IDs, you must use get with the
        appropriate ID to get the details of a Message.
    """
    try:
        response = service.users().messages().list(userId=user_id, q=query).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])
            while 'nextPageToken' in response:
                page_token = response['nextPageToken']
                response = service.users().messages().list(userId=user_id, q=query, pageToken=page_token).execute()
                messages.extend(response['messages'])
        return messages

    except errors.HttpError as error:
        print('[Get message list] Http error occurred: %s' % error)
        return None

def get_attachments(service, msg_id, store_dir, user_id='me'):
    """Get and store attachment from Message with given id.

    Args:
        service: Authorized Gmail API service instance.
        msg_id: ID of Message containing attachment.
        store_dir: The directory used to store attachments.
        user_id: User's email address. The special value "me" can be used to indicate the authenticated user.
    """
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()

        for header in message['payload']['headers']:
            if header['name'] == 'Return-Path':
                sender = header['value'].strip()

        assert(sender != None)
        assert(re.match("(^\<.*\@sfu\.ca\>$)", sender))
        sender = sender[1:-1]

        for part in message['payload']['parts']:
            if part['filename']:
                attachment_id = part['body']['attachmentId']
                attachment = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=attachment_id).execute()
                filename = '/'.join([store_dir, sender, part['filename']])
                filedata = base64.urlsafe_b64decode(attachment['data'].encode('utf-8'))
                os.makedirs(os.path.dirname(filename), exist_ok=True)
                f = open(filename, 'wb')
                f.write(filedata)
                f.close()
                print("📩  ======> Get attachment from mail id %s, saved as %s" % (message['id'], filename))

    except errors.HttpError as error:
        print('[MsgId %s] Http error occurred: %s' % (msg_id, error))
        return None

    except AssertionError:
        print('[MsgId %s] Assertion error occurred, the invalide sender is: %s' % (msg_id, sender))
        return None

    return sender

def set_read(service, msg_id, msg_labels, user_id='me'):
    """Modify the labels on the given message, set it to be read (remove UNREAD label)

    Args:
        service: Authorized Gmail API service instance.
        msg_id: The id of the message required.
        msg_labels: The change in labels.
        user_id: User's email address. The special value "me" can be used to indicate the authenticated user.

    Returns:
        Modified message, containing updated labelIds, id and threadId.
    """
    try:
        message = service.users().messages().modify(userId=user_id, id=msg_id, body=msg_labels).execute()
        label_ids = message['labelIds']
        assert('UNREAD' not in label_ids)

    except errors.HttpError as error:
        print('[MsgId %s] Http error occurred: %s' % (msg_id, error))

    except AssertionError:
        print('[MsgId %s] Assertion error occurred, set read failed.' % msg_id)

def create_message_with_attachment(sender, receiver, subject, message_text, file):
    """Create a message for an email.

    Args:
        sender: Email address of the sender.
        receiver: Email address of the receiver.
        subject: The subject of the email message.
        message_text: The text of the email message.
        file: The path to the file to be attached.

    Returns:
        An object containing a base64url encoded email object.
    """
    message = MIMEMultipart()
    message['to'] = receiver
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    content_type, encoding = mimetypes.guess_type(file)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(file, 'r')
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(file, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(file, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(file, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(file)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_string().encode('UTF-8')).decode('ascii')}

def send_message(service, receiver, message, user_id='me'):
    """Send an email message.

    Args:
        service: Authorized Gmail API service instance.
        receiver: Email address of the receiver.
        message: Message to be sent.
        user_id: User's email address. The special value "me" can be used to indicate the authenticated user.

    Returns:
        Sent Message.
    """
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        msg_id = message['id']
        print("✨  ======> Successfully sent message with id %s to %s" % (msg_id, receiver))
    except errors.HttpError as error:
        print('[MsgId %s] Http error occurred when sending message %s to %s' % (msg_id, to))

def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    print("🔑  ======> Get authorization")

    service = discovery.build('gmail', 'v1', http=http)
    print("📧  ======> Connected to mail service")

    messages = get_list_of_messages(service, query='label:cmpt412 is:unread')
    print("📃  ======> Get unread mail list")

    messages.reverse()
    print(messages)

    for message in messages:
        receiver = get_attachments(service, message['id'], './download')
        if receiver is not None:
            set_read(service, message['id'], {'removeLabelIds': ['UNREAD'], 'addLabelIds': []})
            # do something here for the downloaded attachemnts
            reply = create_message_with_attachment(sender='CMPT412 Bot',
                                                   receiver=receiver,
                                                   subject='[CMPT412] Test Result',
                                                   message_text='blahblah',
                                                   file='./mailbot.py')
            send_message(service, receiver, reply)

if __name__ == '__main__':
    main()
