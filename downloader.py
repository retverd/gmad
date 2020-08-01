import base64
import os.path
import pickle

from apiclient import errors
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

CUR_USER = 'me'
SAVE_FOLDER = './outputs/'
SUBJECT_FILTER = 'Документы по операции № '
TRANS_ID_SIGN = '№ '

# Specifies the name of a file that contains the OAuth 2.0 information for this application,
# including its client_id and client_secret.
CLIENT_SECRETS_FILE = "credentials.json"

# The file token.pickle stores the user's access and refresh tokens, and is created automatically when
# the authorization flow completes for the first time.
TOKEN_DUMP_FILE = 'token.pickle'

# This access scope grants read-only access to the authenticated user's Gmail account.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
API_SERVICE_NAME = 'gmail'
API_VERSION = 'v1'


def get_authenticated_service():
    creds = None

    if os.path.exists(TOKEN_DUMP_FILE):
        with open(TOKEN_DUMP_FILE, 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            with open(CLIENT_SECRETS_FILE, 'rb') as cred:
                print(cred.read())

            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
            creds = flow.run_console()
        # Save the credentials for the next run
        with open(TOKEN_DUMP_FILE, 'wb') as token:
            pickle.dump(creds, token)

    return build(API_SERVICE_NAME, API_VERSION, credentials=creds)


def list_messages_matching_query(service, query: str = ''):
    """List all Messages of the user's mailbox matching the query.

    Args:
      service: Authorized Gmail API service instance.
      query: String used to filter messages returned.
      Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

    Returns:
      List of Messages that match the criteria of the query. Note that the
      returned list contains Message IDs, you must use get with the
      appropriate ID to get the details of a Message.
    """

    # Find among all e-mails first
    try:
        response = service.users().messages().list(userId=CUR_USER, q=query).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=CUR_USER, q=query, pageToken=page_token).execute()
            messages.extend(response['messages'])

        # Check in TRASH too
        response = service.users().messages().list(userId=CUR_USER, q=query, labelIds=['TRASH']).execute()
        if 'messages' in response:
            choice = str(input(f"In Trash bin were found {len(response['messages'])} message(s) matching your request! "
                               f"Would you like to process them too (y/n)? "))
            if choice == 'y':
                messages.extend(response['messages'])
                while 'nextPageToken' in response:
                    page_token = response['nextPageToken']
                    response = service.users().messages().list(userId=CUR_USER, q=query, pageToken=page_token,
                                                               labelIds=['TRASH']).execute()
                    messages.extend(response['messages'])

        return messages
    except errors.HttpError as error:
        print(f'An error occurred: {error}')


def get_attachments(service, user_id: str, msg_id: str, store_dir: str):
    """Get and store attachment from Message with given id.

    Args:
      service: Authorized Gmail API service instance.
      user_id: User's email address. The special value "me"
      can be used to indicate the authenticated user.
      msg_id: ID of Message containing attachment.
      store_dir: The directory used to store attachments.
    """
    trans_id = None

    try:
        # Get message by id
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()

        # Find message subject and get transaction id from it
        for header in message['payload']['headers']:
            if header['name'] == 'Subject':
                _, trans_id = header['value'].split(TRANS_ID_SIGN)

        # Iterate through all message attachments
        for part in message['payload']['parts']:
            # Check if it is attached file
            if part['filename']:
                # Build unique filename
                print(f"Found attachment {part['filename']} for transaction {trans_id}. Labels - {message['labelIds']}")
                file_name = trans_id + '-' + part['filename']

                # Get file content
                if 'data' in part['body']:
                    file_data_dict = part['body']
                elif 'attachmentId' in part['body']:
                    file_data_dict = service.users().messages().attachments().get(userId=user_id, messageId=msg_id,
                                                                                  id=part['body'][
                                                                                      'attachmentId']).execute()
                else:
                    raise RuntimeError(f'Unknown attachment parameters for {part["partId"]} for message {msg_id}')

                # Decode data
                file_data = base64.urlsafe_b64decode(file_data_dict['data'].encode('UTF-8'))

                path = ''.join([store_dir, file_name])

                f = open(path, 'wb')
                f.write(file_data)
                f.close()

    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def main():
    service = get_authenticated_service()

    # Call the Gmail API
    results = list_messages_matching_query(service, query=f'subject:{SUBJECT_FILTER}')

    print(f'Found messages: {len(results)}!')
    # Iterate through list with filtered message ID's and save attachments
    for msg in results:
        get_attachments(service, CUR_USER, msg['id'], SAVE_FOLDER)


if __name__ == '__main__':
    main()
