from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from pandas import read_excel
from datetime import date, datetime, timedelta
from re import match

SCOPES = ['https://www.googleapis.com/auth/calendar']

EVENT_COLORS = {
    'lavender': '1',
    'sage': '2',
    'grape': '3',
    'flamingo': '4',
    'banana': '5',
    'tangerine': '6',
    'peacock': '7',
    'graphite': '8',
    'blueberry': '9',
    'basil': '10',
    'tomato': '11'
}

DAYS_DICT = {'Po': 0, 'Út': 1, 'St': 2, 'Čt': 3, 'Pá': 4}

"""
Uprav funkci custom_format, která regexuje název předmětu 

Pokud chceš upozornění, tak je ve funkci create_event v dict event zakomentovanej blok

Pro lichý a sudý tejdny přičteš 7 dní k proměnný day_od_week_num
    a do event.recurrence přidáš ;INTERVAL=2 (asi na konec, možná za FREQ=WEEKLY)
    
Semestr začíná v podnělí, protože se mi to nechtělo dodělávat
"""

# Začátek a konec semestru
# (YYYY, MM, DD)
START = ('2022', '02', '14')    # první pondělí !
END = ('2022', '05', '16')

FILE_NAME = 'export.xlsx'       # název souboru s exportem rozvrhu
CALENDAR_ID = ''                # id kalendáře; formát: iausd987asdas74554asd@group.calendar.google.com


def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('calendar', 'v3', credentials=creds)

        for event in read_excel_contents(FILE_NAME):
            formated_event = create_event(event)
            res = service.events().insert(calendarId=CALENDAR_ID, body=formated_event).execute()
            print('Event created: %s' % (res.get('htmlLink')))

    except HttpError as error:
        print('An error occurred: %s' % error)


def read_excel_contents(file_name: str, sheet_name='Sheet1'):
    file = read_excel(file_name, sheet_name=sheet_name)
    akce = []
    for index, row in file.iterrows():
        akce.append(row.to_dict())
    return akce


def create_event(event_dict):
    if datetime.strptime(f'{START[0]}-{START[1]}-{START[2]}', '%Y-%m-%d').weekday() != 0:
        raise Exception('Začátek semestru musí začínat pondělím')

    # pro lichý sudý přičteš 7
    day_of_week_num = DAYS_DICT.get(event_dict.get('Den'))  # +7

    actual_day = datetime.strptime(f'{START[0]}-{START[1]}-{START[2]}', '%Y-%m-%d') + timedelta(days=day_of_week_num)
    day_string = date.strftime(actual_day.date(), '%Y-%m-%d')
    color_id = EVENT_COLORS.get('sage') if event_dict.get("Akce") == 'Přednáška' else EVENT_COLORS.get('peacock')

    naz = custom_format(event_dict.get('Předmět'))
    subject_name = naz[1]
    ident = naz[0]

    event = {
        'summary': subject_name,
        'location': f'{event_dict.get("Místnost")}',
        'description': f'{ident}\n{event_dict.get("Vyučující")}',
        'colorId': color_id,
        'start': {
            'dateTime': f'{day_string}T{event_dict.get("Od")}:00',
            'timeZone': 'Europe/Prague',
        },
        'end': {
            'dateTime': f'{day_string}T{event_dict.get("Do")}:00',
            'timeZone': 'Europe/Prague',
        },
        'recurrence': [
            f'RRULE:FREQ=WEEKLY;UNTIL={END[0]}{END[1]}{END[2]}T000000Z'  # pro lichý sůdý přidáš ;INTERVAL=2
        ],
        # 'reminders': {
        #     'useDefault': False,
        #     'overrides': [
        #         {'method': 'email', 'minutes': 24 * 60},
        #         {'method': 'popup', 'minutes': 10},
        #     ],
        # },
    }
    return event


# uprav
def custom_format(event_name: str):
    matched = match(r'([0-9A-Z]+) (.+)', event_name)
    return matched.group(1), matched.group(2)


if __name__ == '__main__':
    main()
