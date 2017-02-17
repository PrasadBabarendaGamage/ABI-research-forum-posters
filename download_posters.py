import os
import httplib2
import json
from datetime import datetime
from apiclient import discovery
from gsuites_api_access import get_drive_credentials, \
                               get_sheets_credentials, \
                               download_file

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

    spreadsheetId = '1rnB9U2I5RN9sZ07b5Olh8MgQYrd70w4UN98ml0C81LM'
    rangeName = 'Form responses 1!A101:L'
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
        poster['id'] = submission[11]
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

        save_pdfs = True
        if save_pdfs:
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