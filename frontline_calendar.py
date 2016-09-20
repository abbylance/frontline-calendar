import argparse
import sys

import arrow
import httplib2
from apiclient import discovery
from appointments import get_credentials, appointments_from_google_sheet, create_google_calendar_events, \
    create_outlook_calendar_events, row_for_name
from exchangelib import DELEGATE
from exchangelib.account import Account
from exchangelib.credentials import Credentials
from oauth2client import tools


def main():
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument('--date', help='What date to start looking at the calendar? Use format YYYY-MM-DD.')
    parser.add_argument('--look_ahead_days', help='How many days to look ahead from the starting date?')
    parser.add_argument('--name', help='Which person are you?')
    parser.add_argument('--google_calendar', action='store_true')
    parser.add_argument('--outlook_calendar', action='store_true')
    parser.add_argument('--spreadsheet_id', help='The ID of the ECBU Luminate Support Weekly Schedule spreadsheet',
                        default='1RgDgDRcyAFDdkEyRH7m_4QOtJ7e-kv324hEWE4JuwgI')
    parser.add_argument('--exchange_username',
                        help='The username you use in Outlook, should be Firstname.Lastname@Blackbaud.me')
    parser.add_argument('--primary_smtp_address',
                        help='Your Outlook email address, should be Firstname.Lastname@blackbaud.com')
    parser.add_argument('--exchange_password', help='The password you use in Outlook')

    flags = parser.parse_args()

    print("Running with args: " + str(sys.argv))

    if not flags.google_calendar or flags.outlook_calendar:
        print("You need to specify --google_calendar and/or --outlook_calendar")
        return

    today = arrow.get(flags.date, 'YYYY-MM-DD')
    dates = [today.replace(days=+n) for n in range(0, int(flags.look_ahead_days))]

    credentials = get_credentials(flags)

    http = credentials.authorize(httplib2.Http())
    sheets_service = discovery.build('sheets', 'v4', http=http)

    google_calendar_service = None
    if flags.google_calendar:
        google_calendar_service = discovery.build('calendar', 'v3', http=http)

    exchange_account = None
    if flags.outlook_calendar:
        exchange_credentials = Credentials(username=flags.exchange_username, password=flags.exchange_password)
        exchange_account = Account(primary_smtp_address=flags.primary_smtp_address, credentials=exchange_credentials,
                                   autodiscover=True, access_type=DELEGATE)

    for date in dates:
        row = row_for_name(sheets_service, flags.spreadsheet_id, flags.name, date)
        if not row:
            print("Could not find row for {name} on {date}, will skip to next day".format(name=flags.name, date=date))
            continue

        midnight = arrow.Arrow(date.year, date.month, date.day, tzinfo='America/Chicago')
        appointments = appointments_from_google_sheet(sheets_service, flags.spreadsheet_id, row, midnight)

        if google_calendar_service:
            events_made = create_google_calendar_events(appointments, google_calendar_service)
            if events_made == 0:
                print("No shifts found for {0}".format(date))

        if exchange_account:
            events_made = create_outlook_calendar_events(appointments, exchange_account)
            if events_made == 0:
                print("No shifts found for {0}".format(date))

if __name__ == "__main__":
    main()
