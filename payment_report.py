"""
Create a new sheet and fill it with payment matching table.
Assumption: The list of balances is in a sheet titled "נוכחי".
"""

# -*- coding: utf-8 -*-

import argparse
import pickle
import os.path
import pprint
import sys
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
spreadsheet_id = '14sq8XzYnEt12S34WKplZUbT9GFa3jOxWlo1fmjyGeIQ'


def main():
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
    spreadsheets = service.spreadsheets()
    # args

    title = 'סילוק חובות עד'
    # CREATE NEW SHEET
    import time
    request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': title,
                        'tabColor': {
                            'red': 1.0,
                            'green': 0.0,
                            'blue': 0.0
                        }
                    }
                }
            }]
        }

    try:
        new_sheet_response = spreadsheets.batchUpdate(spreadsheetId=spreadsheet_id, body=request_body).execute()
    except HttpError as e:
        if b"already exists" in e.content:
            print(f"Sheet named '{title[::-1]}' already exists. Delete the existing one before running.", file=sys.stderr)
            sys.exit()
        raise
    new_sheet_title = new_sheet_response['replies'][0]['addSheet']['properties']['title']
    new_sheet_id = new_sheet_response['replies'][0]['addSheet']['properties']['sheetId']

    balances = get_balances(spreadsheets)
    pprint.pprint(balances)
    payment_matching = determine_payments(balances)
    values, creditor_rows, debtor_row_ranges = format_report(payment_matching)

    # Enter data into the sheet
    request = spreadsheets.values().update(spreadsheetId=spreadsheet_id, range=new_sheet_title + '!A1:ZZ200',
                                           valueInputOption="RAW",
                                           body={'values': values})
    response = request.execute()

    # Set border separators after each creditor
    for row in creditor_rows:
        body = {
            "requests": [
                {
                    "updateBorders": {
                        "range": {
                            "sheetId": new_sheet_id,
                            "startRowIndex": row,
                            "endRowIndex": row + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 6
                        },
                        "bottom": {
                            "style": "SOLID",
                            "width": 1
                        },
                    }
                }
            ]
        }
        res = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    # Set alternating background colors for debtor ranges
    for i, debtor_row_range in enumerate(debtor_row_ranges):
        if i % 2 == 0:
            bgcolor = (208/255.0, 224/255.0, 237/255.0)
        else:
            bgcolor = (245/255.0, 242/255.0, 224/255.0)
        body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": new_sheet_id,
                            "startRowIndex": debtor_row_range[0],
                            "endRowIndex": debtor_row_range[1],
                            "startColumnIndex": 3,
                            "endColumnIndex": 6
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": bgcolor[0],
                                    "green": bgcolor[1],
                                    "blue": bgcolor[2]
                                }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)"
                    }
                }
            ]
        }
        res = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    # Other formattings: general font sizes, merged cells, horizontal alignment, autosize some columns to fit
    body = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": new_sheet_id,
                        "startRowIndex": 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "fontSize": 11
                            }
                        }
                    },
                    "fields": "userEnteredFormat(textFormat)"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": new_sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "fontSize": 16
                            }
                        }
                    },
                    "fields": "userEnteredFormat(textFormat)"
                }
            },
            {
                "mergeCells": {
                    "range": {
                        "sheetId": new_sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 3
                    },
                    "mergeType": "MERGE_ALL"
                }
            },
            {
                "mergeCells": {
                    "range": {
                        "sheetId": new_sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 3,
                        "endColumnIndex": 6
                    },
                    "mergeType": "MERGE_ALL"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": new_sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER"
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment)"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": new_sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 3,
                        "endColumnIndex": 4
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER"
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment)"
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": new_sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 2,
                        "endIndex": 3
                    }
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": new_sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 5,
                        "endIndex": 6
                    }
                }
            }
        ]
    }
    res = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    return


def get_balances(spreadsheets):
    """
    Retrieves list of non-zero balances from sheet titled "נוכחי".
    The returned list structure is [(balance, (name, nick)), ...]
    """
    sheet_metadata = spreadsheets.get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    current_sheet_title = None
    for sheet in sheets:
        title = sheet.get("properties", {}).get("title", "Sheet1")
        if "נוכחי" in title:
            current_sheet_title = title
            break
    if not current_sheet_title:
        print("Sheet with title 'יחכונ' was not found.", file=sys.stderr)
        sys.exit()

    balances = []
    request = spreadsheets.values().get(spreadsheetId=spreadsheet_id, range=current_sheet_title + '!A3:C100')
    response = request.execute()
    values = response.get('values', [])
    for row in values:
        if len(row) >= 3:
            try:
                balance = int(row[0])
            except ValueError:
                balance = None

            # Skip 0 balances and rows where there is no name
            if balance and row[2] and row[2].isascii():
                name = (row[1], row[2])
                balance_entry = (balance, name)
                balances.append(balance_entry)

    # Validate checksum:
    assert sum([x[0] for x in balances]) == 0, "Bad checksum"

    return balances


def determine_payments(balances):
    """
    Matches debtors to creditors, and splits debtor payments as needed.
    The returned list structure is [((creditor balance, (creditor name, nick),
                                                      [(payment value, (debtor name, nick)), ...]), ...]
    """
    debtors = [x for x in balances if x[0] < 0]
    creditors = [x for x in balances if x[0] > 0]

    debtors.sort(key=lambda x: x[0])
    creditors.sort(key=lambda x: x[0], reverse=True)

    # pprint.pprint(debtors)
    # pprint.pprint(creditors)

    payment_matches = []
    # Match naively
    for creditor in creditors:
        debtor_payment_list = []
        owed = creditor[0]
        for i, debtor in enumerate(debtors):
            if abs(debtor[0]) <= owed:
                # Add debtor to list
                debtor_payment_list.append(debtor)
                owed -= abs(debtor[0])
                if owed == 0:
                    break
            else:
                # Add debtor partial payment to list and modify "still to pay" debtor list
                debtor_payment_list.append((-owed, debtor[1]))
                debtors = [(debtor[0] + owed, debtor[1])] + debtors[i+1:]
                break
        payment_matches.append((creditor, debtor_payment_list))
    # pprint.pprint(payment_matches)
    return payment_matches


def format_report(payment_matching):
    """
    Formats 'payment_matching' data as grid values to be updated in the new Google sheet.
    The 'values' structure is a list of rows, each row being a list of cell values.
    Also returns a list of rows where creditors appear and a list of row ranges for debtors.
    These are used for styling the sheet.
    """
    values = [['זכאים', '', '', 'חייבים', '', '']]
    values.append([])
    row_index = 1

    creditor_rows = []
    debtor_row_ranges = []
    last_debtor = payment_matching[0][1][0][1]
    last_debtor_row = 2

    for creditor_match in payment_matching:
        creditor = creditor_match[0]
        first_pass = True  # So that the creditor name and balance appears only once
        for matched_debtor in creditor_match[1]:
            row = [first_pass and creditor[0] or '', first_pass and creditor[1][0] or '', first_pass and creditor[1][1] or '',
                   matched_debtor[0], matched_debtor[1][0], matched_debtor[1][1]]
            row_index += 1
            if matched_debtor[1] != last_debtor:
                debtor_row_range = (last_debtor_row, row_index)
                debtor_row_ranges.append(debtor_row_range)
                last_debtor = matched_debtor[1]
                last_debtor_row = row_index

            first_pass = False
            values.append(row)
        creditor_rows.append(row_index)
        values.append([''])  # empty row
        row_index += 1
    debtor_row_range = (last_debtor_row, row_index)
    debtor_row_ranges.append(debtor_row_range)

    # pprint.pprint(values)
    return values, creditor_rows, debtor_row_ranges


if __name__ == '__main__':
    main()
