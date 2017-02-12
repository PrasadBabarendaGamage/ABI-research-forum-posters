
from __future__ import print_function
import httplib2
import os
import json
from datetime import datetime
#from apiclient import errors

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Drive API Python Quickstart'


def get_drive_credentials():
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
    credential_path = os.path.join(credential_dir,
                                   'drive-python-quickstart.json')

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


def get_sheets_credentials():
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
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

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


def download_file(service, drive_file):
    """Download a file's content.

    Args:
    service: Drive API service instance.
    drive_file: Drive File instance.

  Returns:
    File's content if successful, None otherwise.
  """
    download_url = drive_file.get('downloadUrl')
    if download_url:
        resp, content = service._http.request(download_url)
        if resp.status == 200:
            print('Status: {0}'.format(resp.status))
            return content
        else:
            print('An error occurred: {0}'.format(resp.status))
            return None
    else:
    # The file doesn't have any content stored on Drive.
        return None



def main():
    """ Export ABI research forum posters stored in google drive to a folder.

    The google drive links to the posters are stored in the google sheets that
    is linked to the google form that was sent out to ABI staff and students.
    """
    drive_credentials = get_drive_credentials()
    drive_http = drive_credentials.authorize(httplib2.Http())
    drive_service = discovery.build('drive', 'v2', http=drive_http)

    sheets_credentials = get_sheets_credentials()
    sheets_http = sheets_credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    sheets_service = discovery.build('sheets', 'v4', http=sheets_http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1TArcgKAikMIyc4X70C0yntmbsMeGpG6woLRCqm8FL0I'
    rangeName = 'Form responses 1!A2:K'
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    submissions = result.get('values', [])

    # Create folders for saving poster pdf files.
    datetime.now().replace(second=0, microsecond=0)
    timestamp = str(datetime.now().replace(second=0, microsecond=0))
    all_posters_path = 'all_posters_' + timestamp # for the website and judges
    if not os.path.exists(all_posters_path):
        os.makedirs(all_posters_path)
    print_posters_path = 'posters_to_be_printed_' + timestamp # for IT
    if not os.path.exists(print_posters_path):
        os.makedirs(print_posters_path)

    # Loop through and download each poster submission.
    dict = {}
    dict['posters'] = []
    for submission in submissions:
        poster = {}
        poster['submission time'] = submission[0]
        poster['submitter email address'] = submission[1]
        poster['author'] = submission[2]
        poster['author email address'] = submission[3]
        poster['title'] = submission[4]
        poster['authors string'] = submission[5]
        authors = submission[5].split(',')
        poster['authors list'] = [author.strip() for author in authors]
        poster['abstract'] = submission[6]
        keywords = submission[7].split(',')
        poster['keywords'] = [keyword.strip() for keyword in keywords]
        poster_already_printed = submission[9].split(' (')[0]
        if poster_already_printed == 'Yes':
            poster['print poster'] = False
        else:
            poster['print poster'] = True
        poster['academic status'] = submission[10]
        poster['poster_number'] = ''
        poster['time'] = ''
        poster['url'] = ''

        # Extract google drive poster file id
        file_id = submission[8].split('id=')[-1]
        file = drive_service.files().get(fileId=file_id).execute()
        content = download_file(drive_service, file)
        filename = file['title'].split(' -')[0] + '.pdf'
        filename = filename.replace("{", "")
        filename = filename.replace("}", "")
        poster['filename'] = filename

        # Save file to appropriate folders
        fid = open(os.path.join(all_posters_path,filename), "wb")
        fid.write(content)
        fid.close()
        if poster['print poster']:
            fid = open(os.path.join(print_posters_path, filename), "wb")
            fid.write(content)
            fid.close()
        dict['posters'].append(poster)

    # Save metadata to a json file
    with open('metadata_'+ timestamp + '.json', 'w') as fp:
        json.dump(dict, fp, sort_keys=True, indent=4)

if __name__ == '__main__':
    main()