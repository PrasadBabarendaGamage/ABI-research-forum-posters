import os
import httplib2
import time
import json
import string
from datetime import datetime
from apiclient import discovery
from gsuites_api_access import get_drive_credentials, \
    get_sheets_credentials, \
    download_file


def main():
    """ Sort ABI research forum posters.

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

    spreadsheet_id = '1u3y2l3FAgww4qUcafWtyWIy18mWQ9qk6gg_jE_CMxws'

    # Get submission data and header
    read_rage_name = 'Unsorted!A1:M'
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=read_rage_name).execute()
    submissions = result.get('values', [])

    # Extract header from submission data
    from collections import deque
    submissions = deque(submissions)
    header = submissions.popleft()
    submissions = list(submissions)

    debug = False
    if debug:
        # Only consider the first 5 submissions
        submissions = submissions[0:5]

    # Extract google drive poster file id and add to submission data
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
        poster['poster_number'] = ''
        poster['time'] = ''
        poster['url'] = ''

        if poster['attending']:
            if poster['academic status'] in ['Postgraduate student', 'Post-doc', 'Summer student']:
                # Extract google drive poster file id
                file_id = submission[10].split('id=')[-1]
                file = drive_service.files().get(fileId=file_id).execute()
                #content = download_file(drive_service, file)
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
                submission += [filename]

        # file_id = submission[10].split('id=')[-1]
        # file = drive_service.files().get(fileId=file_id).execute()
        # filename = file['title'].split(' -')[0] + '.pdf'
        # filename = filename.replace("{", "")
        # filename = filename.replace("}", "")
        # submission += [filename]

    # Sort poster submission pdfs in alphabetical order
    from operator import itemgetter
    submissions = sorted(submissions, key=itemgetter(11))

    # Separate summer student submission from others
    summer_student_submissions = []
    other_submissions = []
    other_poster_submissions = []
    pi_submissions = []
    professional_staff_submissions = []
    guest_submissions = []
    for submission in submissions:
        if submission[3] == 'Summer student':
            summer_student_submissions.append(submission)
        elif submission[3] in ['Postgraduate student', 'Post-doc']:
            other_poster_submissions.append(submission)
        elif submission[3] == 'PI':
            pi_submissions.append(submission)
        elif submission[3] == 'Professional Staff':
            professional_staff_submissions.append(submission)
        elif submission[3] == 'Guest':
            guest_submissions.append(submission)
    submissions = summer_student_submissions + other_poster_submissions+pi_submissions+professional_staff_submissions+guest_submissions

    # Add sequential id and alternating poster times to each submission
    idx = 0
    for submission in submissions:
        attending = submission[4]
        if attending == 'Yes':
            if submission[3] in ['Postgraduate student', 'Post-doc', 'Summer student']:
                idx += 1
                if idx % 2 == 0:
                    submission.append('AM')
                else:
                    submission.append('PM')
                submission.append(idx)

    # Add poster filename, time, and id to header
    header += ['poster filename', 'time', 'id']

    # Combine header and submissions
    submissions.insert(0, header)

    # Clear target sheet
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id, range='Sorted!A1:P',
        body={}).execute()

    # Write values to target sheet
    write_range_name = 'Sorted!A1:P'
    body = {
        "majorDimension": "ROWS",
        'values': submissions
    }
    result = sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=write_range_name,
        valueInputOption='USER_ENTERED', body=body).execute()

if __name__ == '__main__':
    main()
