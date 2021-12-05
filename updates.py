from __future__ import print_function
import pandas as pandadata  #data handler
import os # this is to manipulate file deletion for the last cleanup stage.
import datetime # for time stamping
import pickle # google authentication
import os.path  # filesystem utilities
from googleapiclient.discovery import build   # google slide handling
from google_auth_oauthlib.flow import InstalledAppFlow  # google authentication
from google.auth.transport.requests import Request  # google fetch spreadsheet
#import afterprocess
from tabulate import tabulate
from atlassian import Jira
import json
import jmespath as jpath

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '<INSERT GOOGLE SHEET ID HERE>'
SAMPLE_RANGE_NAME = 'Sheet1!A2:A'
#SAMPLE_RANGE_NAME = 'Sheet1'

#pandadata.set_option('display.max_columns', None) #prevents trailing elipses
#pandadata.set_option('display.max_rows', None)
cols = [0,4]  # set the columns

def fetchdata():   #connect to spreadsheet via Google API
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """


    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    #print(values)

    if not values:
        print('No data found.')
    else:
        rows = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                  range=SAMPLE_RANGE_NAME).execute()
        data = rows.get('values')

    return data


def main():

    with open('jiraapitoken.txt', 'r') as myfile:
        myapitokensecret = myfile.readline()

        jira_instance = Jira(
            url = "https://<INSERT_YOURCOMPANY>.atlassian.net/",
            username = "<INSERT YOUR EMAIL>",
            password = myapitokensecret,
        )


    listoftickets = fetchdata()

    mainlist = []

    sublist = ["TicketID", "Summary", "UpdatedBy", "LastComment", "Assignee"]

    mainlist.append(sublist)

    for item in listoftickets:

        sublist = []


        item = str(item).replace('[\'','').replace('\']','').strip()

        #print(item)

        sublist.append(item)

        jiraqueryissue = jira_instance.get_issue(issue_id_or_key= '{}'.format(item))


        summaryquery = jpath.search('fields.summary', jiraqueryissue)

        sublist.append(summaryquery)

        commentqueryid = jpath.search('fields.comment.comments[].id', jiraqueryissue)

        lastcomment = commentqueryid[-1]

        updateAuthorquery = jpath.search("fields.comment.comments[?id=='{0}'].updateAuthor.displayName".format(lastcomment), jiraqueryissue)

        sublist.append(updateAuthorquery)

        #print(updateAuthorquery)

        commentquerybody = jpath.search("fields.comment.comments[?id=='{0}'].body".format(lastcomment), jiraqueryissue)

        #print(commentquerybody)

        sublist.append(commentquerybody)

        #print(sublist)

        assigneequery = jpath.search('fields.assignee.displayName', jiraqueryissue)

        sublist.append(assigneequery)

        mainlist.append(sublist)

    print(tabulate(mainlist, tablefmt="grid"))


    #github test

if __name__ == '__main__':
    main()



