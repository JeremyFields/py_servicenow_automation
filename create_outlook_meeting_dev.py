#!/usr/bin/python
# Script Name: create_outlook_meeting.py
# Purpose: To create a meeting in outlook calendar
# Creation Date: 2022-12-21
# Version: 1.0.0
# Version History: 1.0.0 - Creation
# Author: Jeremy Fields
# *****************************************************************************
# Module Imports

import win32com.client as client
outlook = client.Dispatch("outlook.application") 

# *****************************************************************************
# Global Variables

# *****************************************************************************
# Local Functions

def sendMeeting(date, subject, duration, location, body, organizer):
    appt = outlook.CreateItem(1) # AppointmentItem
    appt.start = date
    appt.subject = subject
    appt.duration = duration # In minutes
    appt.location = location
    appt.body = body
    appt.MeetingStatus = 1 
    appt.Recipients.Add(organizer) 
    appt.OptionalAttendees = "Cloud_Infrastructure_Engineering@xxxx.com"
    appt.display()
    # appt.save()
    # appt.send()

def get_calendar():
    ''' Gets personal calendar from outlook, creates list of chg numbers '''
    chg_num_list = []
    ns = outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items  # for personal calendar
    for app in appointments:
        # if meeting subjects starts with chg number
        if app.subject.startswith("CHG"):
            # get just the chg number from the subject and make list
            subject_split = app.subject.split(" ")
            chg_num = subject_split[0]
            chg_num_list.append(chg_num)

    return chg_num_list

def check_calendar(chg_num_list, chg_num_checker):
    ''' Loops through chg number list, check if chg number from API matches calendar meeting '''
    for chg_num in chg_num_list:
        # if matches, return True (chg is present in calendar)
        if chg_num_checker == chg_num:
            return True
        
    return False

# *****************************************************************************
# Main

def main(chg_num, short_desc, date, duration, organizer):
    ''' creates meetings when passed info from get_changes.py servicenow API script '''
    subject  = f'{chg_num} - {short_desc}'
    body     = f'{short_desc}'
    location = "Current on-call" # Edit
    sendMeeting(date, subject, duration, location, body, organizer)
    
if __name__ == "__main__":
    main()