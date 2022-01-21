# -*- coding: utf-8 -*-

import argparse
import pickle
import os.path
import pprint
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
spreadsheet_id = '14sq8XzYnEt12S34WKplZUbT9GFa3jOxWlo1fmjyGeIQ'

oldname_mapping = {
    'resber777': 'dovlakavod',
    'kavodladov': 'dovlakavod',
    'vodaldovak': 'dovlakavod',
    'mishehu111': 'dovlakavod',
    'JimFridlin': 'connectedJokers',
    'suitedJokers21': 'connectedJokers',
    'arazi100': 'badluck9999',
    'franki19852912': 'franki291285',
    'franki2912': 'franki291285',
    'O.M.G.R2022': 'OMGR888',
    'omrigropper': 'OMGR888',
    'sntran': 'pkrran',
    'talspiderman': 'spidermant57',
    'spidermantal052': 'spidermant57',
    'stam314': 'stam315',
    'thedub11': 'thedub111',
    'THEVERYLUCKYMAN': 'theluckyman3.0',
    'theluckyman2.0': 'theluckyman3.0',
    'ys280981': 'YanivS2809',
    'Yaniv2809': 'YanivS2809',
    'YoavKatz': 'yoavi.k.',
    'elloro99': 'elloro88',
    'Mordillo2020': 'Mordillo2022',
    'Z_707': 'tonnny707',
    'shay830': 'shay4455'
}

def main():
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
    # args
    parser = argparse.ArgumentParser()
    parser.add_argument("-y", "--year", type=str)
    parser.add_argument("-c", "--current", action="store_true")
    args = parser.parse_args()

    # Call the Sheets API
    sheet = service.spreadsheets()

    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    sums = {}
    maxs = {}
    mins = {}
    played_count = {}
    in_money_count= {}
    for sheet in sheets:
        title = sheet.get("properties", {}).get("title", "Sheet1")

        include_current_sheet = args.current and "נוכחי" in title
        include_by_year =  not args.year or args.year in title
        include_regular_sheet = "עד" in title and not "סילוק" in title 
        
        if ((include_regular_sheet and include_by_year) or include_current_sheet):
            sheet_id = sheet.get("properties", {}).get("sheetId", 0)
            # print (title, sheet_id)

            request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=title + '!A3:ZZ100')
            response = request.execute()
            values = response.get('values', [])

            for row in values:
                if len(row) >= 3:
                    try:
                        balance = int(row[0])
                    except ValueError:
                        balance = None
                    username = oldname_mapping.get(row[2]) or row[2]
                    if balance and username and username.isascii():
                        sums[username] = sums.get(username,0) + balance
                        maxs[username] = max(maxs.setdefault(username, 0), balance)
                        mins[username] = min(mins.setdefault(username, 0), balance)
                    if username and row[0]:
                        if username not in played_count:
                            played_count[username] = 0
                        if username not in in_money_count:
                            in_money_count[username] = 0
                        for i in range(0, len(row[3:]), 2):
                            if row[3+i]:
                                played_count[username] += 1
                            try:
                                if row[3+i+1]:
                                    in_money_count[username] += 1
                            except:
                                # hit end
                                pass
    for x in played_count.keys():
        try:
            print (x, '- played:', played_count[x], ', in money:', in_money_count[x],
                   ' - ', "%d%%"%(int(float(in_money_count[x])/played_count[x] * 100.0),))
        except ZeroDivisionError:
            pass
                            

    print ("SUMS:\n")
    sums_l = list(sums.items())
    sums_l.sort(key=lambda x: x[1])
    pprint.pprint(sums_l)
    print ("--------------------------\n")

    print ("MINS:\n")
    mins_l = list(mins.items())
    mins_l.sort(key=lambda x: x[1])
    pprint.pprint(mins_l)
    print ("--------------------------\n")

    print ("MAXS:\n")
    maxs_l = list(maxs.items())
    maxs_l.sort(key=lambda x: x[1])
    pprint.pprint(maxs_l)
    print ("--------------------------\n")

    per_game =  [(sum, played_count[sum[0]], round(float(sum[1])/played_count[sum[0]], 2)) for sum in sums_l]
    per_game.sort(key=lambda x: x[2])
    pprint.pprint(per_game)
                        



if __name__ == '__main__':
    main()
