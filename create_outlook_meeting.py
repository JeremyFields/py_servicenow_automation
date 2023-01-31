# Script Name: create_outlook_meeting.py
# Purpose: To create a meeting in outlook calendar
# Creation Date: 2022-12-21
# Version: 1.0.0
# Version History: 1.0.0 - Creation
# Author: Jeremy Fields
# *****************************************************************************
# Module Imports

import win32com.client as client

# *****************************************************************************
# Global Variables

outlook = client.Dispatch("outlook.application") 

# *****************************************************************************
# Local Functions

def sendMeeting(date, subject, duration, location, body, assigned_to):
    cloud_infra = "Cloud_Infrastructure_Engineering@company.com"
    manager = "manager@company.com"

    appt = outlook.CreateItem(1) # AppointmentItem
    appt.start = date # yyyy-MM-dd hh:mm
    appt.subject = subject
    appt.duration = duration # In minutes
    appt.location = location
    appt.body = body
    appt.MeetingStatus = 1
    appt.BusyStatus = 0
    appt.ResponseRequested = False
    appt.ReminderSet = True
    appt.ReminderMinutesBeforeStart = 120
    #appt.Recipients.Add(assigned_to) -- This is for "required attendees"
    appt.OptionalAttendees = f'{assigned_to}; {manager}; {cloud_infra}'
    appt.display()
    appt.save()
    appt.send()

def get_calendar():
    ''' Gets personal calendar from outlook, creates list of chg numbers '''
    chg_num_list = []
    ns = outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items  # calendar folder
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

def main(chg_num, short_desc, date, duration, assigned_to, location):
    ''' creates meetings when passed info from servicenow API scripts '''
    subject  = f'{chg_num} - {short_desc}'
    body     = f'{short_desc}'
    location = location
    sendMeeting(date, subject, duration, location, body, assigned_to)
    
if __name__ == "__main__":
    main()