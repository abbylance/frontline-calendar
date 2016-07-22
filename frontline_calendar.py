import argparse

import httplib2
import oauth2client
import os
import sys
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
import arrow
from dateutil import tz
from pyexchange import Exchange2010Service, ExchangeNTLMAuthConnection

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


def create_calendar_events(appointments, google_calendar_service, outlook_calendar_service):
    for appointment in appointments:
        summary = None
        if appointment.appointment_type == 'F':
            summary = 'On Phones'
        elif appointment.appointment_type == 'C':
            summary = 'On Chat'

        if summary:
            if not google_calendar_event_exists(appointment, google_calendar_service, summary):
                create_google_calendar_event(appointment, google_calendar_service, summary)
            if not outlook_calendar_event_exists(appointment, outlook_calendar_service, summary):
                create_outlook_calendar_event(appointment, outlook_calendar_service, summary)


def google_calendar_event_exists(appointment, calendar_service, summary):
    matching_events = calendar_service.events().list(calendarId='primary',
                                                     timeMin=appointment.start_time.datetime.isoformat(),
                                                     timeMax=appointment.end_time.datetime.isoformat(),
                                                     q=summary).execute()
    if matching_events and matching_events['items']:
        print("Found {count} matching events for appointment {app}. Will not create a new one.".format(count=len(matching_events['items']), app=appointment))
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
        'description': 'This event was created by Frontline Calendar. Contact charles@connells.org with issues.'
    }

    event = calendar_service.events().insert(calendarId='primary', body=event).execute()
    print('Google Calendar event created. Link: {0} Details: {1}'.format(event.get('htmlLink'), event))


def outlook_calendar_event_exists(appointment, calendar_service, summary):
    matching_events = calendar_service.calendar().list_events(
        start=appointment.start_time.datetime,
        end=appointment.end_time.datetime,
        details=True)

    if matching_events:
        for event in matching_events:
            if event.subject == summary:
                print("Found a matching events for appointment {app}. Will not create a new one.".format(app=appointment))
                return True

    return False


def create_outlook_calendar_event(appointment, calendar_service, summary):
    event = calendar_service.calendar().new_event(
        subject=summary,
        text_body='This event was created by Frontline Calendar. Contact charles@connells.org with issues.',
        start=appointment.start_time.datetime.isoformat(),
        end=appointment.end_time.datetime.isoformat()
    )

    #event.create()
    print('Outlook calendar event created. Details: {0}'.format(event))



def main():
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument('date', help='Which date to examine the ECBU Luminate Support Weekly Schedule for. Use format YYYY-MM-DD.')
    parser.add_argument('row', help='Which row in the spreadsheet is your schedule on?')
    parser.add_argument('--spreadsheet_id', help='The ID of the ECBU Luminate Support Weekly Schedule spreadsheet', default='1RgDgDRcyAFDdkEyRH7m_4QOtJ7e-kv324hEWE4JuwgI')
    parser.add_argument('--exchange_url', help='URL of Blackbaud exchange server, defaults to https://your.email.server.com.here/EWS/Exchange.asmx', default='https://your.email.server.com.here/EWS/Exchange.asmx')
    parser.add_argument('--exchange_username', help='The username you use in Outlook')
    parser.add_argument('--exchange_password', help='The password you use in Outlook')

    flags = parser.parse_args()
    date = arrow.get(flags.date, 'YYYY-MM-DD')

    credentials = get_credentials(flags)
    exe_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    ca_certs = os.path.join(exe_dir, 'certs/cacerts.txt')
    chicago_zoneinfo = os.path.join(exe_dir, 'zoneinfo/Chicago')
    http = credentials.authorize(httplib2.Http(ca_certs=ca_certs))

    sheets_service = discovery.build('sheets', 'v4', http=http)

    # Set up the connection to Exchange
    exchange_connection = ExchangeNTLMAuthConnection(url=flags.exchange_url,
                                                     username=flags.exchange_username,
                                                     password=flags.exchange_password)

    exchange_service = Exchange2010Service(exchange_connection)

    with open(chicago_zoneinfo, 'rb') as zonefile:
        midnight = arrow.Arrow(date.year, date.month, date.day, tzinfo=tz.tzfile(zonefile))
    appointments = appointments_from_google_sheet(sheets_service, flags.spreadsheet_id, flags.row, midnight)

    google_calendar_service = discovery.build('calendar', 'v3', http=http)
    create_calendar_events(appointments, google_calendar_service, exchange_service)

if __name__ == "__main__":
    main()
