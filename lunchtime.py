import argparse
import sys
from math import floor

import arrow
import httplib2
from apiclient import discovery
from appointments import get_credentials, appointments_from_google_sheet, create_google_calendar_events, \
    create_outlook_calendar_events, Range, Appointment, LUNCH, row_for_name
from exchangelib import DELEGATE
from exchangelib.account import Account
from exchangelib.credentials import Credentials
from oauth2client import tools

LUNCH_EARLIEST = 11.0
LUNCH_LATEST = 16.0


def fractional_hour(datetime):
    return datetime.hour + (float(datetime.minute) / 60)


def hour_minute_from_fractional_hour(fractional):
    return int(floor(fractional)), int((fractional - floor(fractional)) * 60)


def main():
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument('--date', help='What date to start looking at the calendar? Use format YYYY-MM-DD.')
    parser.add_argument('--look_ahead_days', help='How many days to look ahead from the starting date?')
    parser.add_argument('--first_name', help='Name of first person on lunch date')
    parser.add_argument('--second_name', help='Name of second person on lunch date')
    parser.add_argument('--google_calendar', action='store_true')
    parser.add_argument('--outlook_calendar', action='store_true')
    parser.add_argument('--spreadsheet_id', help='The ID of the ECBU Luminate Support Weekly Schedule spreadsheet', default='1RgDgDRcyAFDdkEyRH7m_4QOtJ7e-kv324hEWE4JuwgI')
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
        midnight = arrow.Arrow(date.year, date.month, date.day, tzinfo='America/Chicago')

        first_row = row_for_name(sheets_service, flags.spreadsheet_id, flags.first_name, date)
        if not first_row:
            print("Could not find row for {name} on {date}, will skip to next day".format(name=flags.first_name, date=date))
            continue

        second_row = row_for_name(sheets_service, flags.spreadsheet_id, flags.second_name, date)
        if not second_row:
            print("Could not find row for {name} on {date}, will skip to next day".format(name=flags.second_name, date=date))
            continue

        first_appointments = appointments_from_google_sheet(sheets_service, flags.spreadsheet_id, first_row, midnight)
        second_appointments = appointments_from_google_sheet(sheets_service, flags.spreadsheet_id, second_row, midnight)

        if date.weekday() in [5, 6]:  # skip weekends
            continue

        # if no appointments are found, don't try to schedule lunch
        if not first_appointments or not second_appointments:
            print("No schedule yet defined for {0}".format(date))
            continue

        lunch_ranges = [Range(LUNCH_EARLIEST, LUNCH_LATEST)]
        for appointment in first_appointments + second_appointments:
            if appointment.appointment_type in ['F', 'C', 'PTO']:
                new_lunch_ranges = []
                appointment_range = Range(fractional_hour(appointment.start_time), fractional_hour(appointment.end_time))
                for lunch_range in lunch_ranges:
                    new_lunch_ranges.extend(lunch_range.subtract(appointment_range))
                lunch_ranges = new_lunch_ranges

        # cut out ranges that are too short
        lunch_ranges = [r for r in lunch_ranges if r.length() >= 1.5]

        lunch_appointments = [Appointment(midnight.replace(hour=hour_minute_from_fractional_hour(r.start)[0],
                                                           minute=hour_minute_from_fractional_hour(r.start)[1]),
                                          midnight.replace(hour=hour_minute_from_fractional_hour(r.end)[0],
                                                           minute=hour_minute_from_fractional_hour(r.end)[1]),
                                          LUNCH) for r in lunch_ranges]

        if google_calendar_service:
            events_made = create_google_calendar_events(lunch_appointments, google_calendar_service)
            if events_made == 0:
                print("No shifts found for {0}".format(date))

        if exchange_account:
            events_made = create_outlook_calendar_events(lunch_appointments, exchange_account)
            if events_made == 0:
                print("No shifts found for {0}".format(date))

if __name__ == "__main__":
    main()
