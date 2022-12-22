# py_servicenow_automation

Some of these scripts are standalone and some will only work in tandem with others as they import functions from those other scripts. <br>
I will edit this list to detail any script dependencies. <br>

# Standalone:

# Dependent:
### ---change request automation---
    Main script     : get_changes_dev.py    --> Main script imports and utilizes both of the imported scripts.  
    Imported scripts: create_outlook_meeting.py  
                    : check_calendar.py
