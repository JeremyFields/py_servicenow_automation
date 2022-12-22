#!/usr/bin/python
"""NAME"""
# Script Name: check_calendar.py
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
chg_num_list = []

# *****************************************************************************
# Local Functions

def get_calendar():
    ''' Gets personal calendar from outlook, creates list of chg numbers '''
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

def main(chg_num_checker):
    ''' Main called from servicenow API script, gets passed chg num from servicenow table '''
    chg_num_list = get_calendar()
    is_present = check_calendar(chg_num_list, chg_num_checker)
    # returns whether chg number is present in calendar or not to servicenow API script
    return is_present

if __name__ == "__main__":
    main()
        