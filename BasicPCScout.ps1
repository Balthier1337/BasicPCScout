# 
# Author: Joseph McConville
# Website: http://www.lionheartservices.co.uk
# Created: 2015-04-01
# Modified: 2015-04-02
# Script: PC Information Collector
# Version: 1.0
# Description: Collects various information from the PC, and outputs to a CSV file.
# 

function Voodoo {
	# Collect the info from WMI
	$computerSystem = get-wmiobject Win32_ComputerSystem
	$computerBIOS = get-wmiobject Win32_BIOS
	$computerOS = get-wmiobject Win32_OperatingSystem
	$computerCPU = get-wmiobject Win32_Processor
	$computerHDD = Get-WmiObject Win32_LogicalDisk -Filter drivetype=3
	# Create the variables containing filtered data
	$PCName = $computerSystem.Name
	$Manufacturer = $computerSystem.Manufacturer
	$Model = $computerSystem.Model
	$SerialNumber = $computerBIOS.SerialNumber
	$RAM = "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB)
	$CPU = $computerCPU.Name
	$OS = $computerOS.caption
	$SP = $computerOS.ServicePackMajorVersion
	$User = $computerSystem.UserName
	$BootTime = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
}

#Run the function to populate the PC variables
Voodoo

##################################
##### Create custom settings #####
##################################
$csvFolder = [environment]::getfolderpath('mydocuments')
$csvFile = "\system-info.csv"

# Extra variables
$csv = "$csvFolder$csvFile"
$csvCheck = Test-Path $csv
$dupeCheck = if ($csvCheck) {Get-Content $csv | Select-String $PCName}

# If PC hasn't already been placed in the file, add it
if (-Not ($dupeCheck)) { 
	# If the CSV file has already been created, add a new row
	if ($csvCheck) {
		'{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}' -f $PCName,$Manufacturer,$Model,$SerialNumber,$RAM,$CPU,$OS,$SP,$User,$BootTime | Add-Content -path $csv
	}
	# If the CSV file doesn't exist, create it, add the headers, and add the PC
	elseif (-Not ($csvCheck)) {
		'{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}' -f "PCName","Manufacturer","Model","SerialNumber","RAM","CPU","OS","SP","User","BootTime" | Add-Content -path $csv
		'{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}' -f $PCName,$Manufacturer,$Model,$SerialNumber,$RAM,$CPU,$OS,$SP,$User,$BootTime | Add-Content -path $csv
	}
}
