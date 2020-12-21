# Epic Printers
This was built to manage a pool of EPS servers.  It will create (including multiple trays and configuration), deletes, and shows the status of print queues.

.DESCRIPTION
This script should be run as an administrator in an elevated Powershell session.
This script uses powershell, WMI, and PrintUI.exe to create print queues on all your EPS servers and apply settings based on the tray type and driver.  
To do this, the script needs to create and maintain binary files specific to tray types and drivers.  Please store this script in a central location where EPS admins have write access.
A folder will be created called PrinterBin and binary files will be stored there.  Using a standard naming convention, the script will look for a binary file to apply the correct settings
to a print queue.  If the file doesn't exist, the queue will be created and the print configuration window will be opened and the user will be prompted to configure the queue the correct way.
A binary file will be exported for this driver and tray type for use with future queues.

This script can be run with or without parameters.  If run without parameters, a menu will walk the user through creating print queues one printer at a time. 

.PARAMETER Function
REQUIRED - The function parameter allows the user to perform functions from the command line and pass the necessary information without going through the menu.  
Valid options:  
    Create - Creates a print queue using other parameters, if the required parameters aren't specified, it will ask 
    Delete - Deletes queues specified in parameter -PrinterName
    Status - Displays the network status and queue info for all queues that start with the queues specified in -PrinterName
All functions require the parameter -PrinterName to also be specified.

.PARAMETER PrinterName
REQUIRED - This is the name of the printer you want to work with.  For the delete function, you can specify a partial name if you want to delete several queues that begin with same characters

.PARAMETER Environment 
This allows you change between the production and non production environments.  The default is the production environment.  To switch to the nonproduction environment, enter: -Environment NP

.PARAMETER Trays
When creating printers, this selects which trays to build.  This should be an integer value that corresponds to the selection from the menu.

.PARMETER Driver
When creating printers, this selects which driver to use.  This should be an integer value that corresponds to the selection from the menu.

.PARAMETER Comment
When creating printers, this specifies the value for the comment field.  If $AskforComment=0 in settings, $DefaultComment will be used and this is obsolete.

.PARAMETER Location
When creating printers, this specifies the value for the location field.  If $AskforLocation=0 in settings, $DefaultLocation will be used and this is obsolete.

.EXAMPLE
EpicPrinters.ps1 -Function Create -PrinterName Printer1 -Trays 1 -Driver 0 -Comment "Created By PowerShell Script" -Location "Admitting Front Desk"

.EXAMPLE
EpicPrinters.ps1 -Function Delete -PrinterName Printer

.EXAMPLE
EpicPrinters.ps1 -Function Status -PrinterName Printer2

.NOTES
Author: Mitch Duff, me@mitchduff.com, 317-509-7681 (contact me with questions or suggestions)
#>