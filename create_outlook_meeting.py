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

def sendMeeting(date, subject, duration, location, body):
    appt = outlook.CreateItem(1) # AppointmentItem
    appt.start = date # yyyy-MM-dd hh:mm
    appt.subject = subject
    appt.duration = duration # In minutes
    appt.location = location
    appt.body = body
    appt.MeetingStatus = 1 
    appt.Recipients.Add('jeremy.fields@xxxx.com') 
    appt.display()
    appt.save()
    appt.send()


# *****************************************************************************
# Main

def main(chg_num, short_desc, date, duration, location):
    ''' creates meetings when passed info from get_changes.py servicenow API script '''
    subject  = f'test {chg_num} - {short_desc}'
    body     = f'{short_desc}'
    print(chg_num, short_desc, date, duration, location)
    sendMeeting(date, subject, duration, location, body)
    
if __name__ == "__main__":
    main()