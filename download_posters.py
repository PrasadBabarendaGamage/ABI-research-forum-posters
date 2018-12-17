import os
import httplib2
import json
import string
from datetime import datetime
from apiclient import discovery
from gsuites_api_access import get_drive_credentials, \
                               get_sheets_credentials, \
                               download_file


def list_duplicates(seq):
    seen = set()
    seen_add = seen.add
    return [idx for idx,item in enumerate(seq) if item in seen or seen_add(item)]

def find_duplicates(submissions):
    email_address = []
    for submission in submissions:
        email_address.append(submission[1])

    for index in list_duplicates(email_address):
        print submissions[index][1]

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
    sorted_spreadsheet = True
    if sorted_spreadsheet:
        spreadsheetId = '1u3y2l3FAgww4qUcafWtyWIy18mWQ9qk6gg_jE_CMxws'  # Sorted
        rangeName = 'Sorted!A2:P116'
    else:
        spreadsheetId = '188rGjYaLF0C5UfVDzs5QInvtH1V7HQ6efugpJpDNqCo' # Original
        rangeName = 'Form responses 1!A143:M'

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    submissions = result.get('values', [])

    # Create folders for saving poster pdf files.
    datetime.now().replace(second=0, microsecond=0)
    timestamp = str(datetime.now().replace(second=0, microsecond=0))
    all_posters_path = 'all_posters_' + timestamp # for the website and judges
    if not os.path.exists(all_posters_path):
        os.makedirs(all_posters_path)
    if not sorted_spreadsheet:  # Only for printing
        print_non_summer_posters_path = 'non_summer_posters_to_be_printed_' + timestamp # for IT
        print_summer_posters_path = 'summer_posters_to_be_printed_' + timestamp # for IT
        if not os.path.exists(print_non_summer_posters_path):
            os.makedirs(print_non_summer_posters_path)
        if not os.path.exists(print_summer_posters_path):
            os.makedirs(print_summer_posters_path)

    find_duplicates(submissions)

    # Loop through and download each poster submission.
    dict = {}
    dict['posters'] = []
    for submission in submissions:
        poster = {}
        poster['submission time'] = submission[0]
        poster['author email address'] = submission[1]
        poster['author'] = submission[2]
        poster['academic status'] = submission[3]
        attending = submission[4]
        if attending == 'Yes':
            poster['attending'] = True
        else:
            poster['attending'] = False
        poster['reason for not attending'] = submission[5]
        poster['title'] = submission[6]
        poster['authors string'] = submission[7]
        authors = submission[7].split(',')
        poster['authors list'] = [author.strip() for author in authors]
        poster['abstract'] = submission[8]
        keywords = submission[9].split(',')
        poster['keywords'] = [keyword.strip() for keyword in keywords]
        poster_already_printed = submission[11].split(' (')[0]
        if poster_already_printed == 'Yes':
            poster['print poster'] = False
        else:
            poster['print poster'] = True
        poster['Dietary requirements'] = submission[12]

        if poster['attending']:
            if poster['academic status'] in ['Postgraduate student', 'Post-doc', 'Summer student']:
                # Extract google drive poster file id
                file_id = submission[10].split('id=')[-1]
                file = drive_service.files().get(fileId=file_id).execute()
                content = download_file(drive_service, file)

                if sorted_spreadsheet: # Only for website
                    poster['filename'] = submission[13]
                    poster['time'] = submission[14]
                    poster['poster_number'] = submission[15]
                    poster['url'] = ''
                else: # Only for printing
                    postfix = ''
                    filename = file['title'].split(' -')[0] + postfix
                    if 'abir' in filename:
                        filename = string.replace(filename, '2017', '2018', 1)
                        filename = filename.replace("{", "")
                        filename = filename.replace("}", "")
                    else:
                        filename = poster['author'] + '_abirf2018.pdf'
                    if not filename.endswith('.pdf'):
                        filename += '.pdf'
                    print filename
                    poster['filename'] = filename

                save_pdfs = True
                if save_pdfs:
                    # Save file to appropriate folders
                    fid = open(os.path.join(all_posters_path,poster['filename']), "wb")
                    fid.write(content)
                    fid.close()
                    if not sorted_spreadsheet: # Only for printing
                        if poster['print poster']:
                            if poster['academic status'] == 'Summer student':
                                poster_path = print_summer_posters_path
                            else:
                                poster_path = print_non_summer_posters_path
                            fid = open(os.path.join(poster_path, poster['filename']), "wb")
                            fid.write(content)
                            fid.close()

        dict['posters'].append(poster)

    # Save metadata to a json file
    with open('metadata_'+ timestamp + '.json', 'w') as fp:
        json.dump(dict, fp, sort_keys=True, indent=4)

if __name__ == '__main__':
    main()