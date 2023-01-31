# Script Name    : get_change_details.py
# Purpose        : automate meeting creations for servicenow changes
# Creation Date  : 2023-01-31
# Version        : 1.0.0
# Version History: 1.0.0 - Creation
# Author         : Jeremy Fields
# *****************************************************************************
# Module Imports

import pysnow, pysnow.exceptions
from datetime import datetime, timedelta
import create_outlook_meeting_dev
from pprint import pprint
import json

# *****************************************************************************
# Instances

## PROD
instance = 'xxx'
password = 'xxx'
user = 'xxx'

## DEV
# instance = 'xxx'
# user = 'xxx'
# password = 'xxx'

# *****************************************************************************
# Global Variables

client = pysnow.Client(instance=instance, user=user, password=password, raise_on_empty=True)
change_task_table = client.resource(api_path='/table/change_task')
user_group_table = client.resource(api_path='/table/sys_user_group')
change_request_table = client.resource(api_path='/table/change_request')
user_table = client.resource(api_path='/table/sys_user')

today = datetime.today()
two_days_ago = today - timedelta(days=2)
qb = (
        pysnow.QueryBuilder()
        .field('active').equals('true')
        .AND()
        .field('sys_created_on').between(two_days_ago, today)
    )

# *****************************************************************************
# Local Functions

def get_user_group_id():
    groups = ['AWS Cloud Infrastructure Engineering', 'Infrastructure: Unix', 'Infrastructure: Middleware']
    sys_ids = {}
    for group in groups:
        response = user_group_table.get(query={'name': group}, stream=True)
        for response in response.all():
            sys_ids[group] = response["sys_id"]
    
    return sys_ids

def get_user_id(change_request_and_task_list):
    final_cr_list = []
    for cr_and_task in change_request_and_task_list:
    
        qb = (
            pysnow.QueryBuilder()
            .field('active').equals('true')
            .AND()
            .field('sys_id').equals(cr_and_task["Assigned To"])
        )

        response = user_table.get(query=qb, stream=True)
        for record in response.all():
            cr_and_task["Assigned To"] = record["name"]
            cr_and_task["Location"] = f'{record["name"]}\'s desk'
            cr_and_task["Email"] = record["email"]
            final_cr_list.append(cr_and_task)
    
    return final_cr_list


def get_change_task(sys_id_dict):
    list_of_change_tasks = []
    response = change_task_table.get(query=qb, stream=True)
    for response in response.all():
        
        assignment_group = response["assignment_group"]
        for key in assignment_group:
            if key == 'value':
                for k, v in sys_id_dict.items():
                    if sys_id_dict[k] == assignment_group[key]:
                        change_task_dict = {}
                        change_task_dict["Team"] = k
                        change_task_dict["Task Number"] = response["number"]
                        change_task_dict["Expected Start"] = response["expected_start"]
                        change_task_dict["Task Description"] = response["short_description"]
                        change_task_dict["Change Request"] = response["change_request"]["value"]
                        change_task_dict["Assigned To"] = response["assigned_to"]["value"]

                        list_of_change_tasks.append(change_task_dict)

    return list_of_change_tasks
                        
def compare_tasks_to_requests(list_of_change_tasks):
    change_request_and_task_list = []
    qb = (
        pysnow.QueryBuilder()
        .field('active').equals('true')
        .AND()
        .field('sys_created_on').between(two_days_ago, today)
        .AND()
        .field('state').equals('-2')
    )
    response = change_request_table.get(query=qb, stream=True)
    for response in response.all():
        for task in list_of_change_tasks:

            if task["Change Request"] == response["sys_id"]:
                task["Change Request"] = response["number"]
                task["Change Description"] = response["short_description"]
                task["Start Date"] = response["start_date"]
                task["End Date"] = response["end_date"]
                change_request_and_task_list.append(task)
    
    return change_request_and_task_list
            
def parse_data_set_meetings(final_cr_list):
    change_tracker = []     # To make sure multiple meetings for same CHG (separate TASKS) doesn't occur.
    for change_request in final_cr_list:
        if change_request["Change Request"] not in change_tracker:
            change_tracker.append(change_request["Change Request"])
            print(f'\n{change_request["Change Request"]} NOT ACCOUNTED FOR\n')
            start_date = datetime.strptime(change_request["Start Date"], "%Y-%m-%d %H:%M:%S")
            end_date = datetime.strptime(change_request["End Date"], "%Y-%m-%d %H:%M:%S")
            # calculate duration of chg
            duration = calculate_duration(start_date, end_date)
            start_date, end_date = calculate_time(start_date, end_date)
            chgs_on_calendar = create_outlook_meeting_dev.get_calendar()
            is_present = create_outlook_meeting_dev.check_calendar(chgs_on_calendar, change_request["Change Request"])
            # if it's not present, create a meeting
            if not is_present:
                print(f'Creating meeting for {change_request["Change Request"]} which will begin on {start_date}.')
                create_outlook_meeting_dev.main(change_request["Change Request"], change_request["Change Description"], 
                                                start_date, duration, change_request["Email"], change_request["Location"])
                # EXPORT TO FILE
                create_file_archive(change_request)
                
            else:
                print(f'\n{change_request["Change Request"]} is present in calendar. Skipping meeting creation.\n')
        else:
            print(f'Different task: {change_request["Task Number"]} - {change_request["Change Request"]} already accounted for.')

def calculate_duration(start_date, end_date):
    ''' calculates meeting duration for outlook '''
    duration = end_date - start_date
    # get duration in seconds
    duration = duration.total_seconds()
    # convert to minutes for outlook
    duration = duration / 60

    return duration

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

def create_file_archive(change_dict):
    ''' exports change requests to file '''
    with open('C:/Users/Jeremy.Fields/Documents/Scripts/Python/servicenow/change_tsk_request_archive_dev.txt', 'a') as f:
        f.write("Meeting created for: ")
        f.write(json.dumps(change_dict))
        f.write("\n")

# *****************************************************************************
# Main

def main():
    # Get user groups (AWS CIE, Middleware, Unix)
    sys_group_id_dict = get_user_group_id()
    # get change tasks assigned to the user groups
    list_of_change_tasks = get_change_task(sys_group_id_dict)
    # adding CR numbers to dict for the change tasks
    change_request_and_task_list = compare_tasks_to_requests(list_of_change_tasks)
    # Final list - converting user sys_id to friendly names, add to dict
    final_cr_list = get_user_id(change_request_and_task_list)
    # set meetings
    parse_data_set_meetings(final_cr_list)

if __name__ == "__main__":
    main()