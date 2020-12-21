<#
.SYNOPSIS
The script creates (including multiple trays and configuration), deletes, and shows the status of printers on EPS servers.

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
Param([string]$Function,[string]$PrinterName,[string]$Environment,[string]$Trays,[string]$Driver,[string]$Comment,[string]$Location)
#####################################################
function CreatePrinter{
    Param([string]$PrinterName,[string]$Driver,[string]$Trays,[string]$Comment,[string]$Location)
    ShowText -Text "Build Printers" -Color Yellow
    If($PrinterName -eq ""){$PrinterName=AskaQuestion -Question "Enter the name of the printer"}
    $PrinterName=$PrinterName.ToUpper()
    If($UseIPAddress -eq 1){$PortName=AskaQuestion -Question "Enter the IP address of the printer"}Else{$PortName=$PrinterName}
    If($AskForComment -eq 1){If($Comment -eq ""){$Comment=AskAQuestion -Question "Enter Comments"}}Else{$Comment=$DefaultComment}
    If($AskForLocation -eq 1){If($Location -eq ""){$Location=AskAQuestion -Question "Enter Comments"}}Else{$Location=$DefaultLocation}
    If(($Driver.ToString() -eq -1) -or ($Driver.ToString() -eq "")){
        Showtext -Text "Select the Driver" -Color Cyan
        $i=0
        ForEach($PrintDriver in $PrintDriverNames){ShowText -Text ($i.ToString()+". "+$PrintDriver);$i++}
        $Selection=AskaQuestion -Question "Selection" -BlankOK 0
        $DriverName=$PrintDriverNames[$Selection]}
    Else{$DriverName=$PrintDriverNames[$Driver]}
    #$Trays=0
    If($Trays.ToString() -eq ""){
        ShowText -Text "Select the trays to build" -Color Cyan
        $i=0
        ForEach($PrinterTrayName in $PrinterTrayNames){ShowText -Text ($i.ToString()+". "+$PrinterTrayName);$i++}
        $Selection=AskaQuestion -Question "Selection" -BlankOK 0
        $TraysToBuild=$PrinterTrayNames[$Selection]}
    Else{$TraysToBuild=$PrinterTrayNames[$Trays]}
    $TrayNames=$TraysToBuild.Split("|")
    ShowText -Text ("Creating printer - "+$PrinterName) -color Yellow
    ForEach($Server in $EPSServers){
        CreatePort -Server $Server -HostName $PortName
        ForEach($TrayName in $TrayNames){
            $TrayName=$TrayName.Replace("Plain","")
            CreateQueue -Server $Server -DriverName $DriverName -Printer ($PrinterName+$TrayName) -Port $PortName -Comment $Comment -Location $Location
            SetTrayDefault -Server $Server -DriverName $DriverName -Printer ($PrinterName+$TrayName) -TrayName $TrayName}}
    ShowText -Text "printer creation complete" -Color Yellow
    $PrinterName=""
    $DriverName=""
    $TrayNames=$Null
    $TrayToBuild=""
    $Driver=""
    $Trays="0"
    $Selection=""
}
#####################################################
function DeletePrinter{
    Param([string]$PrinterName)
    ShowText -Text ("Deleting Printer - "+$PrinterName) -Color "Yellow"
    $AllTrayNames=@()
    ForEach($TrayName in $PrinterTrayNames){
        $TrayName=$TrayName.Replace("Plain","")
        $TrayName=$TrayName.Replace("(Beaker)","")
        $AllTrayNames+=$TrayName.Split("|")
    }
    $AllTrayNames=$AllTrayNames | sort -unique
    ForEach($Server in $EPSServers){
        ForEach($TrayName in $AllTrayNames){
            ShowText -Text ("Deleting print queue: \\"+$Server+"\"+($PrinterName+$TrayName)) -Color Green
            DeleteQueue -Server $Server -Printer ($PrinterName+$TrayName)
        }}}
#####################################################
function CreateQueue{
    Param ([string]$Server,[string]$DriverName,[string]$Printer,[string]$Port,[string]$Comment,[string]$Location)
	$Printer=$Printer.Replace("(Beaker)","")
    ShowText -Text ("creating print queue: \\"+$Server+"\"+$Printer+" with "+$DriverName) -Color Green
    $print=([WMICLASS]"\\$Server\ROOT\cimv2:Win32_Printer").createInstance()
	$print.DriverName=$DriverName
    $print.PortName=$Port
    $print.Shared=$False
	$print.DeviceID=$Printer 
    $print.Comment=$Comment
	$print.Location=$Location
	$print.RawOnly = $True
	$print.EnableBIDI= $False
    $print.DoCompleteFirst=$True  
	$print.Put()|Out-Null
    $print=$Null
}
#####################################################
function DeleteQueue{
    Param ([string]$Server,[string]$Printer)
    $print=([WMICLASS]"\\$Server\ROOT\cimv2:Win32_Printer").CreateInstance()
	$print.DeviceID=$Printer
	$print.CancelAllJobs()|Out-Null
    $print.Delete()|Out-Null
    $print=$Null
}
#####################################################
function CreatePort{
    Param([string]$Server,[string]$HostName)
    ShowText -Text ("creating printer port: '"+$Hostname+"' on " +$Server) -Color Green
    $port=([WMICLASS]"\\$Server\ROOT\cimv2:Win32_TCPIPPrinterPort").createInstance() 
    $port.Name=$HostName
    $port.SNMPEnabled=$False 
    $port.Protocol=1 
    $port.HostAddress=$HostName
    $port.Put()|Out-Null
    $port=$Null
}
#####################################################
function SetTrayDefault{
    Param([string]$Server,[string]$DriverName,[string]$Printer,[string]$TrayName)
    $Printer=$Printer.Replace("(Beaker)","")
    ShowText -Text ("setting printer defaults: \\"+$Server+"\"+$Printer) -Color Green
    $BinFile=$PSScriptRoot+"\PrinterBin\"+$DriverName+$TrayName+'.bin'
    If(Test-Path $BinFile){SetBinFile -Server $Server -Printer $Printer -BinFile $BinFile}Else{CreateBinFile -Server $Server -Printer $Printer -BinFile $BinFile}
}
#####################################################
function CreateBinFile{
    Param([string]$Server,[string]$Printer,[string]$BinFile)
    PrintUI.exe /n \\$Server\$Printer /p
    ShowText -Text "Please configure queue defaults for this tray type, save and close the window."
	If((AskaQuestion -Question "Type done when ready to export printer settings") -eq "done"){
        PrintUI.exe /Ss /n \\$Server\$Printer /a $BinFile
        ShowText -Text ("creating bin file: "+$BinFile) -Color "Green"
        Start-Sleep -s 5}
}
#####################################################
function SetBinFile{
    Param([string]$Server,[string]$Printer,[string]$Binfile)
    Start-Sleep 6
    PrintUI.exe /Sr /n \\$Server\$Printer /q /a $BinFile r d u g
}
#####################################################
#This section requires customization for each model used.  The code downloads the web page for the printer and tries to determine the model by finding unique text.
#There is probably a better way to do this but i haven't found.
Function GetPrinterModel{ 
    param([string] $PrinterName) 
    $wc = New-Object System.Net.WebClient 
    [string]$Output = ($wc.downloadstring("http://"+$PrinterName)).ToUpper()
    If(($Output.Contains("HP") -or $Output.Contains("WCD/INDEX.HTML"))){Return 0}
    If($Output.Contains("DELL")){ Return 3}
    If($Output.Contains("EASYCODER")){Return 2}
    $Output = ($wc.downloadstring("http://"+$PrinterName+"/status/status.htm")).ToUpper()
    If($Output.Contains("DELL")){Return 3}
    If($Output.Contains("M5")){Return 4}
    $wc = New-Object System.Net.WebClient 
    $Output = ($wc.downloadstring("http://"+$PrinterName+"/index.lua")).ToUpper()
    If($Output.Contains("PM43")){Return 1}
    $Output = ($wc.downloadstring("http://"+$PrinterName+"/hp/jetdirect")).ToUpper()
    If($Output.Contains("HP")){Return 0}
    If($Output -eq $Null){Return 0}
}
#####################################################
function ShowPrinterStatus{
	Param ([string]$PrinterName)
    $Filter="Name like'"+$PrinterName+"%'"
    ShowText -Text ("Printer Status - "+$PrinterName) -Color Yellow
    ShowText -Text "Network Status: " -Color Cyan -NoNewLine 1
    $Status=ShowStatus -Type "Network" -Name $PrinterName
    Write-Host
    ShowText -Text "Printer Model: " -Color Cyan -NoNewLine 1
    If($Status -eq "ONLINE"){
        $Driver=GetPrinterModel $PrinterName
        If($Driver -eq -1){$PrinterDriver="Unknown (Check Model)"}Else{$PrinterDriver=$PrintDriverNames[$Driver]}}
    Else{$PrinterDriver="Unknown (Offline)"}
    Write-Host $PrinterDriver
    ShowText -Text ((AddTab -Text "server" -Space 20)+(AddTab -Text "Queue" -Space 15)+(AddTab -Text "Driver" -Space 40)+(AddTab -Text "Status" -Space 10)+(AddTab -Text "# Jobs")) -Color Cyan
	ForEach($Server in $EPSServers){
        $JobStatus = Get-WMIObject Win32_PerfFormattedData_Spooler_PrintQueue -ComputerName $Server -Filter $Filter | Select Jobs,Name
        If($JobStatus -eq $Null){
            ShowText -Text (AddTab -Text $Server -Space 30) -NoNewLine 1
            ShowText -Text "No Queue" -Color Red}
        Else{
            ForEach($Queue in $JobStatus){
                $QueueStatus=([WMICLASS]"\\$Server\ROOT\cimv2:Win32_Printer").CreateInstance()
                $QueueStatus.DeviceID=$Queue.Name
                $QueueStatus.Get()
                ShowText -Text (AddTab -Text $Server -Space 20) -NoNewLine 1
                ShowText -Text (AddTab -Text $Queue.Name -Space 15) -NoNewLine 1
                ShowText -Text (AddTab -Text $QueueStatus.DriverName -Space 40) -NoNewLine 1
                $Status=ShowStatus -Type "Queue" -Name $QueueStatus.PrinterStatus -Parameter $QueueStatus.ExtendedPrinterStatus -AddTab 10
                $Status=ShowStatus -Type "Count" -Name $Queue.Jobs -AddTab 10
 
               Write-Host}}}
    Return $Driver 
 }
#####################################################
function ShowStatus{
    Param([string]$Name,[string]$Type,[int]$AddTab,[string]$Parameter )
    Switch($Type){
        "Network" {If(($Online=GetNetworkStatus -ComputerName $Name) -eq 0){$Color="Green"}Else{$Color="Red"};$Status=$NetworkStatus[$Online]}
        "Count"  {$Status=$Name;If($Name -eq 0){$Color="Green"}Else{$Color="Red"}}
        "Queue"   {If($PrinterStatus[$Name] -eq "Other"){$Status=$ExPrinterStatus[$Parameter];$Color="Red"}Else{$Status=$PrinterStatus[$Name];$Color="Green"}}}
    If($AddTab -gt 0){$DisplayStatus=AddTab -Text $Status -Space $AddTab}Else{$DisplayStatus=$Status}
    ShowText -Text $DisplayStatus -Color $Color -NoNewLine 1
    Return $Status
}
#####################################################
function GetNetworkStatus{
	Param([string]$ComputerName)
    If($ComputerName -eq ""){$ComputerName=AskaQuestion "Enter the HostName to check"}
        If((Test-Connection -ComputerName $ComputerName -count 2 -quiet) -eq $False){
		    Try{$Records=[Net.DNS]::GetHostEntry($ComputerName)
            If($Records.HostName -ne $Null){Return 1} }
	Catch{Return 2}}
    Else{Return 0}
}
#####################################################
function AddTab{
    Param([string]$Text,[int]$Space)
    $Char=$Space-$Text.length
    $AfterTab=$Text
	For($i=1;$i -le $Char;$i++){$AfterTab+=" "}
	Return $AfterTab
}
#####################################################
function AskaQuestion{
	Param([string]$Question,[int]$BlankOK)
    ShowText -Text ($Question+": ") -Color Cyan -NoNewLine 1
	$Answer=Read-Host
	While(($Answer -eq "")-and ($BlankOK -ne 1)){
        ShowText -Text "RESPONSE CANNOT BE BLANK" -Color Red
        ShowText -Text ($Question+": ") -Color Cyan -NoNewLine 1
        $Answer=Read-Host}
	Return $Answer
}
#####################################################
function ShowText{
    Param([string]$Text,[string]$Color="White",[int]$NoNewLine)
    If($NoNewLine -eq 1){Write-Host -foreground $Color -NoNewLine $Text.ToUpper()}Else{Write-Host -foreground $Color $Text.ToUpper()}
}
#####################################################
function GetDrivers{
    $FullDrivers = Get-WmiObject -ComputerName $ProdServers[0] -Query "select Name from Win32_PrinterDriver"
    $DriverList=@()
    ForEach($Driver in $FullDrivers){$SplitName=$Driver.Name.Split(",");$DriverList+=$SplitName[0]}
    Return $DriverList
}  

#####################################################
function Pause{Read-Host 'Press Enter to continue...' | Out-Null}
#####################################################
$ErrorActionPreference='silentlycontinue'
#####################################################
#SETTINGS
#Update this section with your environment
#####################################################
#Exact Name of Drivers to use with the script
$PrintDriverNames=@("HP UNIVERSAL PRINTING PCL 5 (V5.9.0)")
#Characters that are added to the printer name to designate each tray type, use Plain if the tray type adds nothing to printer name.  The script will remove "Plain" before creating the queues.
$PrinterTrayNames=@("Plain","RX","Plain|RX")
#Change to 1 if the port for the queue uses the IP address
$UseIPAddress=0
#Change to 1 if the script should ask for the comment field
$AskForComment=0
#Change to 1 if the script should ask for the location field
$AskForLocation=0
#If AskforComment=0, this value is enter in the comment field
$DefaultComment="Created by Powershell Script"
#If AskforLocation=0, this value is entered in the location field
$DefaultLocation=""
#Enter the non-prod EPS servers
$NonProdServers=@("testserver1","testserver2")
#Enter the prod EPS servers
$ProdServers=@("prodserver1","prodserver2","prodserver3","prodserver4","prodserver5","prodserver6")
#####################################################
#####################################################
$NetworkStatus=@("ONLINE","OFFLINE","NO DNS")
$PrinterStatus=@("","Other","Unknown","Ready","Printing","Warming Up","Stopped Printing","Offline")
$ExPrinterStatus=@("","Other","Unknown","Idle","Printing","Warming Up","Stopped Printing","Offline","Paused","Error","Busy","Not Available","Waiting","Processing","Initialization","Power Save","Pending Deletion","I/O Active","Manual Feed")
If($Environment -eq "NP"){$EPSServers=$NonProdServers}Else{$EPSServers=$ProdServers}
If($Function -ne ""){
    If($PrinterName -eq ""){ShowText -Text "No Printer Identified" - color "Red";Exit}
    Switch ($Function){
        "Create" {$Driver=GetPrinterModel $PrinterName;CreatePrinter -PrinterName $PrinterName -Driver $Driver -Trays $Trays -Comment $Comment -Location $Location}
        "Delete" {DeletePrinter -PrinterName $PrinterName}
        "Status" {ShowPrinterStatus -PrinterName $PrinterName}
        default {ShowText -Text "No function specified" -Color "Red"}}
    Exit}
While (1){
    $PrinterName=AskAQuestion "Enter the Name of the Printer"
    $Driver=ShowPrinterStatus -PrinterName $PrinterName
    $Answer=AskAQuestion "Type Yes to build/rebuild printer, anything else to continue"
    If($Answer -eq "yes"){CreatePrinter -PrinterName $PrinterName -Driver $Driver -Trays ""}
    Pause
    Clear-Host}