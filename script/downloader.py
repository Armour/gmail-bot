from __future__ import print_function
import httplib2
import os
import base64

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
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
    user_id: User's email address. The default special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

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
    print('An error occurred: %s' % error)

def get_attachments(service, msg_id, store_dir, user_id='me'):
  """Get and store attachment from Message with given id.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: ID of Message containing attachment.
    store_dir: The directory used to store attachments.
  """
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()

    sender = message['snippet'].strip().split()[0]

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
    print('An error occurred: %s' % error)


def main():
    """Download attached files for each student """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    print("🔑  ======> Get authorization")

    service = discovery.build('gmail', 'v1', http=http)
    print("📧  ======> Connected to mail service")

    messages = get_list_of_messages(service, query='label:cmpt412')
    print("📃  ======> Get mail list")

    for message in messages:
        get_attachments(service, message['id'], './download')

if __name__ == '__main__':
    main()
