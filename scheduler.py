#!/usr/bin/env python3

#
# Nick Feamster
# July 2020
# Use Google calendar API to produce simple text output
#

from __future__ import print_function

import time
from datetime import datetime, timezone, timedelta

from dateutil.parser import *
from dateutil.tz import *
from dateutil.relativedelta import *

from datetimerange import DateTimeRange

import json
import argparse

import pickle
import sys
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

########
# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

# Simple spec for time format
datefmt = '%a %b %d'
timefmt = '%H:%M'
dtfmt = datefmt + ' ' + timefmt

########

def free_busy(service):

    # round the time down to the previous half-hour from now
    timeMin = datetime.utcnow()
    timeMin = timeMin - timedelta(minutes=timeMin.minute % 30, \
                                    seconds=timeMin.second, \
                                    microseconds=timeMin.microsecond)

    timeMin = timeMin.isoformat() + 'Z' # 'Z' indicates UTC time

    # look a full week ahead
    finweekdate = datetime.utcnow() + relativedelta(days=+15)
    timeMax = finweekdate.isoformat() + 'Z'


    print ('\n=== Available Times ({}) ===\n'.format(time.localtime().tm_zone))

    # Fix: query multiple calendars
    calendar_id = 'primary'
    body = {
        "timeMin": timeMin,
        "timeMax": timeMax,
        "timeZone": 'UTC',
        "items": [{"id": calendar_id}]
    }

    freebusy_result = service.freebusy().query(body=body).execute()
    slots = freebusy_result.get('calendars', {}).get(calendar_id, [])
    
    if not slots:
        print('No busy slots in the specified timeframe.')

    # get a list of the busy intervals from Google Calendar
    busy_times = []
    for busy in slots['busy']:
        (dts, dte) = (parse(busy['start']), parse(busy['end']))
        busy_times.append((dts,dte))

        # Print the busy slots (debug)
        # print("{} - {}".format(dts.astimezone(tzlocal()), dte.astimezone(tzlocal())))


    last_b = 1
    # intersect each possible time range with the busy slots
    # test each 30-minute timeslot within the date range
    time_range = DateTimeRange(timeMin, timeMax)
    for value in time_range.range(relativedelta(minutes=+30)): 
            # busy varuable
            b = 0

            # fix: make this configurable
            # block out times except for those within a range of the day
            local_time = value.astimezone(tzlocal())
            if local_time.hour < 13 or \
                    local_time.hour > 17 or \
                    local_time.weekday() > 4:
                b = 1
            
            slot_min = value + relativedelta(minutes=+1)
            slot_max = value + relativedelta(minutes=+29)
            slot_range = DateTimeRange(slot_min, slot_max)

            for busy in slots['busy']:
                # busy range
                tr = DateTimeRange(busy['start'],busy['end'])
                x = DateTimeRange(slot_min, slot_max)

                if tr.is_intersection(x):
                    b = 1

            if last_b and not b:
                start_free = local_time + relativedelta(minutes=-0)

            if not last_b and b:
                print(start_free.strftime(dtfmt), '-', local_time.strftime(timefmt))
            last_b = b

########

def print_today(service):


    t = datetime.utcnow()
    local_time = t.astimezone(tzlocal())

    timeMin = datetime(t.year, t.month, t.day)
    timeMin = timeMin.isoformat() + 'Z' # 'Z' indicates UTC time
    enddate = datetime(t.year, t.month, t.day, 23, 59, 59)
    timeMax = enddate.isoformat() + 'Z'

    events_result = service.events().list(calendarId='primary', timeMin=timeMin, timeMax=timeMax, singleEvents=True, orderBy='startTime').execute()
    events = events_result.get('items', [])

    print('\n= {} =\n= Contents =\n'.format(local_time.strftime('%Y-%m-%d')))

    if not events:
        print('No upcoming events found.')

    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        starttime = parse(start)

        # Vimwiki Format
        print('== {} ({}) =='.format(event['summary'], starttime.strftime(timefmt)))

########

def print_week(service):

    print ('\n=== Week Schedule ===\n')
    now = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    finweekdate = datetime.now() + relativedelta(weeks=+1)
    finweek = finweekdate.isoformat() + 'Z'

    # Call the Calendar API to get the week's events
    events_result = service.events().list(calendarId='primary', timeMin=now, timeMax=finweek, singleEvents=True, orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        starttime = parse(start)
        print(starttime.strftime(dtfmt), '\t', event['summary'])

########

def print_next(service):
    now = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time

    print('=== Next 10 Events ===')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=10, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        starttime = parse(start)
        print(starttime.strftime(dtfmt), event['summary'])


#######
def get_creds():

    creds = None
    path = os.path.dirname(os.path.realpath(__file__))

    cfile = path + '/credentials.json'
    tfile = path + '/token.pickle'

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists(tfile):
        with open(tfile, 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                cfile, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(tfile, 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service

########

def main():
    
    # Argument Parsing
    parser = argparse.ArgumentParser(description='Parse Schedule.')
    parser.add_argument('-f', '--free', help='print free times', action="store_true") 
    parser.add_argument('-n', '--next', help='print next 10 appointments', action="store_true") 
    parser.add_argument('-w', '--week', help='print next week of appointments', action="store_true") 
    args = parser.parse_args()

    # Set up API
    service = get_creds()

    # Print various outputs
    if len(sys.argv) < 2:
        print_today(service)

    if args.week:
        print_week(service)

    if args.next:
        print_next(service)

    if args.free:
        free_busy(service)
    
if __name__ == '__main__':
    main()
