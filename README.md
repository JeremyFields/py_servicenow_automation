# py_servicenow_automation

Description:
Script that queries change tasks within the last X days (currently set to 2) that are assigned to
AWS Cloud Infrastructure Engineering, Infrastructure: Middleware, Infrastructure: Unix.
Creates a dictionary with the system ID's for those groups. Takes the Change Request table link value
and queries the change request table in order to match the change task to the change request number and
add to the dictionary. Queries the sys_user table to add the "assigned to" user to the dictionary.
Lastly, calls the "create_outlook_meeting.py" script to create meetings in outlook for the 
change requests that were found to be assigned to the groups. <br>

Libraries needed:
pysnow==0.7.17
pywin32==305
from datetime import datetime, timedelta
import create_outlook_meeting
import json

### ---change request automation---
    Main script     : get_change_details.py    --> Main script imports and utilizes both of the imported scripts.  
    Imported script: create_outlook_meeting.py  

