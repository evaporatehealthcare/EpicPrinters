# Epic Printers
This was built to manage a pool of EPS servers.  It will create (including multiple trays and configuration), deletes, and shows the status of print queues.

# DESCRIPTION
This script should be run as an administrator in an elevated Powershell session.
This script uses powershell, WMI, and PrintUI.exe to create print queues on all your EPS servers and apply settings based on the tray type and driver.  
To do this, the script needs to create and maintain binary files specific to tray types and drivers.  Please store this script in a central location where EPS admins have write access.
A folder will be created called PrinterBin and binary files will be stored there.  Using a standard naming convention, the script will look for a binary file to apply the correct settings
to a print queue.  If the file doesn't exist, the queue will be created and the print configuration window will be opened and the user will be prompted to configure the queue the correct way.
A binary file will be exported for this driver and tray type for use with future queues.

This script can be run with or without parameters.  If run without parameters, a menu will walk the user through creating print queues one printer at a time. 


Author: 
Mitch Duff
mitch@evaporatehealthcare.com
317-509-7681 
(contact me with questions or suggestions)