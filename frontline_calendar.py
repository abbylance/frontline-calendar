import argparse
from datetime import datetime

import httplib2
import oauth2client
import os
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
import arrow
from exchangelib import DELEGATE, AllProperties
from exchangelib.credentials import Credentials
from exchangelib.account import Account
from exchangelib.ewsdatetime import EWSDateTime, EWSTimeZone
from exchangelib.folders import CalendarItem

SCOPES = ['https://www.googleapis.com/auth/calendar',
          'https://www.googleapis.com/auth/spreadsheets.readonly']

CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Frontline Calendar'

FIRST_CELL_MINUTES_AFTER_MIDNIGHT = 7 * 60


class Appointment:
    start_time = 0
    end_time = 0
    appointment_type = None

    def __repr__(self):
        return "{0}: {1} - {2}".format(self.appointment_type, self.start_time, self.end_time)


def get_credentials(flags):
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'drive-python-frontline-calendar.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        flags.noauth_local_webserver = True
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials


def time_from_cell_index(index, day):
    minutes = FIRST_CELL_MINUTES_AFTER_MIDNIGHT + (index * 15)
    return day.replace(minutes=minutes)


def appointments_from_google_sheet(service, spreadsheet_id, row, midnight):
    rangeName = "'{date}'!K{row}:BF{row}".format(row=row, date=midnight.strftime("%a %m.%d.%y"))

    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=rangeName).execute()
    time_blocks = result.get('values', [])
    # shed outer list
    if isinstance(time_blocks, list):
        time_blocks = time_blocks[0]

    appointments = []
    currentAppointment = None
    for i in range(0, len(time_blocks)):
        time_block_type = time_blocks[i]
        # if there is an appointment in progress of the same type, extend it
        if currentAppointment and currentAppointment.appointment_type == time_block_type:
            currentAppointment.end_time = time_from_cell_index(i+1, midnight)
        # start a new appointment
        else:
            if currentAppointment:
                appointments.append(currentAppointment)
            currentAppointment = Appointment()
            currentAppointment.start_time = time_from_cell_index(i, midnight)
            currentAppointment.end_time = time_from_cell_index(i+1, midnight)
            currentAppointment.appointment_type = time_block_type

    # clean up the last appointment
    appointments.append(currentAppointment)

    return appointments


def create_google_calendar_events(appointments, google_calendar_service):
    for appointment in appointments:
        summary = None
        if appointment.appointment_type == 'F':
            summary = 'On Phones'
        elif appointment.appointment_type == 'C':
            summary = 'On Chat'

        if summary and not google_calendar_event_exists(appointment, google_calendar_service, summary):
            create_google_calendar_event(appointment, google_calendar_service, summary)


def google_calendar_event_exists(appointment, calendar_service, summary):
    matching_events = calendar_service.events().list(calendarId='primary',
                                                     timeMin=appointment.start_time.datetime.isoformat(),
                                                     timeMax=appointment.end_time.datetime.isoformat(),
                                                     q=summary).execute()
    if matching_events and matching_events['items']:
        print("Found {count} matching Google Calendar events for appointment {app}. Will not create a new one.".format(count=len(matching_events['items']), app=appointment))
        return True

    return False


def create_google_calendar_event(appointment, calendar_service, summary):
    event = {
        'summary': summary,
            'start': {
                'dateTime': appointment.start_time.datetime.isoformat()
            },
            'end': {
                'dateTime': appointment.end_time.datetime.isoformat()
            },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {
                    'method': 'popup',
                    'minutes': '5'
                },
                {
                    'method': 'popup',
                    'minutes': '0'
                }
            ]
        },
        'description': 'This event was created by Frontline Calendar. Contact charles@connells.org with issues.'
    }

    event = calendar_service.events().insert(calendarId='primary', body=event).execute()
    print('Google Calendar event created. Link: {0} Details: {1}'.format(event.get('htmlLink'), event))


def create_outlook_calendar_events(appointments, outlook_calendar_service):
    for appointment in appointments:
        summary = None
        if appointment.appointment_type == 'F':
            summary = 'On Phones'
        elif appointment.appointment_type == 'C':
            summary = 'On Chat'

        if summary and not outlook_calendar_event_exists(appointment, outlook_calendar_service, summary):
            create_outlook_calendar_event(appointment, outlook_calendar_service, summary)


def outlook_calendar_event_exists(appointment, calendar_service, summary):
    ews_tz = EWSTimeZone.timezone('America/Chicago')

    d = appointment.start_time.datetime
    start_date_time = datetime(d.year, d.month, d.day, d.hour, d.minute, d.second, d.microsecond, ews_tz)
    start_ews_date_time = EWSDateTime.from_datetime(start_date_time)

    d = appointment.end_time.datetime
    end_date_time = datetime(d.year, d.month, d.day, d.hour, d.minute, d.second, d.microsecond, ews_tz)
    end_ews_date_time = EWSDateTime.from_datetime(end_date_time)

    matching_events = calendar_service.calendar.find_items(
        start=start_ews_date_time,
        end=end_ews_date_time,
        shape=AllProperties)

    if matching_events:
        for event in matching_events:
            if event.subject == summary:
                print("Found a matching Outlook calendar event for appointment {app}. Will not create a new one.".format(app=appointment))
                return True

    return False


def create_outlook_calendar_event(appointment, calendar_service, summary):
    ews_tz = EWSTimeZone.timezone('America/Chicago')

    d = appointment.start_time.datetime
    start_date_time = datetime(d.year, d.month, d.day, d.hour, d.minute, d.second, d.microsecond, ews_tz)
    start_ews_date_time = EWSDateTime.from_datetime(start_date_time)

    d = appointment.end_time.datetime
    end_date_time = datetime(d.year, d.month, d.day, d.hour, d.minute, d.second, d.microsecond, ews_tz)
    end_ews_date_time = EWSDateTime.from_datetime(end_date_time)

    event = CalendarItem(
        subject=summary,
        body='This event was created by Frontline Calendar. Contact charles@connells.org with issues.',
        start=start_ews_date_time,
        end=end_ews_date_time
    )

    calendar_service.calendar.add_items([event])
    print('Outlook calendar event created. Details: {0}'.format(event))


def main():
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument('date', help='Which date to examine the ECBU Luminate Support Weekly Schedule for. Use format YYYY-MM-DD.')
    parser.add_argument('row', help='Which row in the spreadsheet is your schedule on?')
    parser.add_argument('--google_calendar', action='store_true')
    parser.add_argument('--outlook_calendar', action='store_true')
    parser.add_argument('--spreadsheet_id', help='The ID of the ECBU Luminate Support Weekly Schedule spreadsheet', default='1RgDgDRcyAFDdkEyRH7m_4QOtJ7e-kv324hEWE4JuwgI')
    parser.add_argument('--exchange_username', help='The username you use in Outlook, should be Firstname.Lastname@Blackbaud.me')
    parser.add_argument('--exchange_password', help='The password you use in Outlook')

    flags = parser.parse_args()

    if not flags.google_calendar or flags.outlook_calendar:
        print("You need to specify --google_calendar and/or --outlook_calendar")
        return

    date = arrow.get(flags.date, 'YYYY-MM-DD')

    credentials = get_credentials(flags)

    http = credentials.authorize(httplib2.Http())
    sheets_service = discovery.build('sheets', 'v4', http=http)
    midnight = arrow.Arrow(date.year, date.month, date.day, tzinfo='America/Chicago')
    appointments = appointments_from_google_sheet(sheets_service, flags.spreadsheet_id, flags.row, midnight)

    if flags.google_calendar:
        google_calendar_service = discovery.build('calendar', 'v3', http=http)
        create_google_calendar_events(appointments, google_calendar_service)

    if flags.outlook_calendar:
        exchange_credentials = Credentials(username=flags.exchange_username, password=flags.exchange_password)
        exchange_account = Account(primary_smtp_address='Abigail.Lance@blackbaud.com', credentials=exchange_credentials, autodiscover=True, access_type=DELEGATE)
        create_outlook_calendar_events(appointments, exchange_account)

if __name__ == "__main__":
    main()
