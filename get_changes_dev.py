#!/usr/bin/python
# Script Name    : get_changes.py
# Purpose        : Automation script to get the scheduled change requests 
#                  from servicenow and creates meetings in outlook.
# Creation Date  : 2022-12-21
# Version        : 1.0.0
# Version History: 1.0.0 - Creation
# Author         : Jeremy Fields

''' TO DO:
    ONLY Get changes assigned to our team
    If assigned to AWS Cloud Infra? Unix?
        '''

# *****************************************************************************
# Module Imports

import pysnow, pysnow.exceptions
from datetime import datetime, timedelta
import create_outlook_meeting
import argparse
import json
from pprint import pprint

# *****************************************************************************
# Argparser

parser = argparse.ArgumentParser()
parser.add_argument("--organizer", default="Cloud_Infrastructure_Engineering@xxxx.com")
args = parser.parse_args()

# *****************************************************************************
# Global Variables

# DEV
instance = 'snowautomation'
user = 'infrastructue_automation'
password = 'xxxx'
client = pysnow.Client(instance=instance, user=user, password=password, raise_on_empty=True)
change_table = client.resource(api_path='/table/change_request')


# *****************************************************************************
# Local Functions

def get_scheduled_CR():
    # query CR table for changes that are in scheduled state (state: -2)
    response = change_table.get(query={'state': -2}, stream=True)
    
    return response

def parse_data_set_meetings(response):
    ''' Parses through scheduled chgs, creates meetings if needed '''
    chg_num_list = []
    for record in response.all():
        # pprint(record)
        if record['u_change_owner']['value'] == '12345678901024858549ffdrhf4': # or record['u_change_owner_group'] == "AWS Cloud Infrastructure Engineering" ??
        # create dictionary of info from scheduled chg 
            change_dict = {
                "change_number": record['number'],
                "opened_at": record['opened_at'],
                "start_date": record['start_date'],
                "end_date": record['end_date'],
                "opened_at": record['opened_at'],
                "location": record['location'],
                "description": record['description'],
                "short_description": record['short_description'],
                "change_owner": "Technical Program Manager"
            }

            pprint(change_dict)

            start_date = datetime.strptime(change_dict['start_date'], "%Y-%m-%d %H:%M:%S")
            end_date = datetime.strptime(change_dict['end_date'], "%Y-%m-%d %H:%M:%S")
            # calculate duration of chg
            duration = calculate_duration(start_date, end_date)
            start_date, end_date = calculate_time(start_date, end_date)
            # get current changes on the calendar
            chg_num_list = create_outlook_meeting.get_calendar()
            # check if chg from servicenow API is already present on the calendar
            is_present = create_outlook_meeting.check_calendar(chg_num_list, change_dict['change_number'])
            # if it's not present, create a meeting
            if not is_present:
                print(f"Creating meeting for {change_dict['change_number']} which will begin on {start_date}.")
                create_outlook_meeting.main(change_dict['change_number'], change_dict['short_description'], \
                                    start_date, duration, args.organizer)
                # EXPORT TO FILE
                create_file_archive(change_dict)
                
            else:
                print(f"{change_dict['change_number']} is present in calendar. Skipping meeting creation.")

def calculate_time(start_date, end_date):
    ''' changes GMT to GMT-7 and returns timestamp '''
    # Convert to GMT-7 full timestamp
    start_date = start_date - timedelta(hours=7)
    end_date = end_date - timedelta(hours=7)
    # get time only in 12 hour format
    start_time = start_date.strftime("%I:%M:%S %p")
    end_time = end_date.strftime("%I:%M:%S %p")
    # get the date only
    start_date = start_date.date()
    end_date = end_date.date()
    # combine to make full 12 hour timestamp in GMT-7
    start_date = f"{start_date} {start_time}"
    end_date = f"{end_date} {end_time}"

    return start_date, end_date
    
def calculate_duration(start_date, end_date):
    ''' calculates meeting duration for outlook '''
    duration = end_date - start_date
    # get duration in seconds
    duration = duration.total_seconds()
    # convert to minutes for outlook
    duration = duration / 60

    return duration

def create_file_archive(change_dict):
    ''' exports change requests to file '''
    with open('change_request_archive.txt', 'a') as f:
        f.write("Meeting created for: ")
        f.write(json.dumps(change_dict))
        f.write("\n")

    
# *****************************************************************************
# Main

def main():
    response = get_scheduled_CR()
    parse_data_set_meetings(response)

if __name__ == "__main__":
    main()