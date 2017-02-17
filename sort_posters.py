import os
import httplib2
import time
import json
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

    spreadsheet_id = '1rnB9U2I5RN9sZ07b5Olh8MgQYrd70w4UN98ml0C81LM'

    # Get submission data and header
    read_rage_name = 'Form responses 1!A1:K'
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
        file_id = submission[8].split('id=')[-1]
        file = drive_service.files().get(fileId=file_id).execute()
        filename = file['title'].split(' -')[0] + '.pdf'
        filename = filename.replace("{", "")
        filename = filename.replace("}", "")
        submission += [filename]

    # Sort poster submission pdfs in alphabetical order
    from operator import itemgetter
    submissions = sorted(submissions, key=itemgetter(11))

    # Separate summer student submission from others
    summer_student_submissions = []
    other_submissions = []
    for submission in submissions:
        if submission[10] == 'Summer student':
            summer_student_submissions.append(submission)
        else:
            other_submissions.append(submission)
    submissions = summer_student_submissions + other_submissions

    # Add sequential id and alternating poster times to each submission
    for idx, submission in enumerate(submissions):
        if idx % 2 == 0:
            submission.append('AM')
        else:
            submission.append('PM')
        submission.append(idx+1)

    # Add poster filename, time, and id to header
    header += ['poster filename', 'time', 'id']

    # Combine header and submissions
    submissions.insert(0, header)

    # Clear target sheet
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id, range='Sorted!A1:Z',
        body={}).execute()

    # Write values to target sheet
    write_range_name = 'Sorted!A1:N'
    body = {
        "majorDimension": "ROWS",
        'values': submissions
    }
    result = sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=write_range_name,
        valueInputOption='USER_ENTERED', body=body).execute()

if __name__ == '__main__':
    main()
