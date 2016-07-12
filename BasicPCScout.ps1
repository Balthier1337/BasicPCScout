# 
# Author: Joseph McConville
# Website: http://www.lionheartservices.co.uk
# Created: 2015-04-01
# Modified: 2015-04-02
# Script: PC Information Collector
# Version: 2.0
# Description: Collects various information from the PC, and outputs to a CSV file.
###################################################################################

##################################
##### Create custom settings #####
##################################
$csvFolder = ""
$csvFile = "\system-info.csv"
$xlsFile = "\system-info.xls"

##################################
##### Do Not Edit After Here #####
##################################

#
# Collect the info from WMI
#
$computerSystem = get-wmiobject Win32_ComputerSystem
$computerBIOS = get-wmiobject Win32_BIOS
$computerOS = get-wmiobject Win32_OperatingSystem
$computerCPU = get-wmiobject Win32_Processor
$computerHDD = Get-WmiObject Win32_LogicalDisk -Filter drivetype=3
$MSOffice = Get-WmiObject -class Win32_product | where {$_.Description -Like 'Microsoft Office*' -And $_.Description -NotLike '*MUI*' -And $_.Description -NotLike '*Components*' -And $_.Description -NotLike '*Runtime*' -And $_.Description -NotLike '*Proof*'}
$win32os = Get-WmiObject Win32_OperatingSystem -computer "."

#
# Create the variables containing filtered data
#
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
$Office = $MSOffice.Description
$Arch = $win32os.OSArchitecture
$ID = $win32os.SerialNumber
#
# Extra variables
#
$csv = "$csvFolder$csvFile"
$xlsFolder = $csvFolder
$xls = "$xlsFolder$xlsFile"
$csv = "$csvFolder$csvFile"
$csvCheck = Test-Path $csv
$xlsCheck = Test-Path $xls
$dupeCheck = if ($csvCheck) {Get-Content $csv | Select-String $PCName}

#
# Functions
#

function addPC() {
	# If PC hasn't already been placed in the file, add it
	if (-Not ($dupeCheck)) { 
		# If the CSV file has already been created, add a new row
		if ($csvCheck) {
			'{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}' -f $PCName,$Manufacturer,$Model,$SerialNumber,$RAM,$CPU,$Arch,$OS,$SP,$ID,$Key,$Office,$OfficeKey,$OfficeID,$User,$BootTime | Add-Content -path $csv
		}
		# If the CSV file doesn't exist, create it, add the headers, and add the PC
		elseif (-Not ($csvCheck)) {
			'{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}' -f "PCName","Manufacturer","Model","SerialNumber","RAM","CPU","Architecture","OS","SP","OS UID","Windows Key","Office","Office Key","Office ID","User","BootTime" | Add-Content -path $csv
			'{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}' -f $PCName,$Manufacturer,$Model,$SerialNumber,$RAM,$CPU,$Arch,$OS,$SP,$ID,$Key,$Office,$OfficeKey,$OfficeID,$User,$BootTime | Add-Content -path $csv
		}
	}
}

function varPCTasks() {
	powercfg -h off
}

function Search-RegistryKeyValues {
	param(
	[string]$path,
	[string]$valueName
	)
	Get-ChildItem $path -recurse -ea SilentlyContinue | 
	% { 
		if ((Get-ItemProperty -Path $_.PsPath -ea SilentlyContinue) -match $valueName)
		{
			$_.PsPath
		} 
	}
}

function getOSVersion() {
    param ($targets = ".")
    $hklm = 2147483650
    $regPath = "Software\Microsoft\Windows NT\CurrentVersion"
    $regValue = "DigitalProductId"
    Foreach ($target in $targets) {
        $productKey = $null
        $win32os = $null
        $wmi = [WMIClass]"\\$target\root\default:stdRegProv"
        $data = $wmi.GetBinaryValue($hklm,$regPath,$regValue)
        $binArray = ($data.uValue)[52..66]
        $charsArray = "B","C","D","F","G","H","J","K","M","P","Q","R","T","V","W","X","Y","2","3","4","6","7","8","9"
        ## decrypt base24 encoded binary data
        For ($i = 24; $i -ge 0; $i--) {
            $k = 0
            For ($j = 14; $j -ge 0; $j--) {
                $k = $k * 256 -bxor $binArray[$j]
                $binArray[$j] = [math]::truncate($k / 24)
                $k = $k % 24
			}
            $productKey = $charsArray[$k] + $productKey
            If (($i % 5 -eq 0) -and ($i -ne 0)) {
                $productKey = "-" + $productKey
			}
		}
		Return $productKey
	}
}

function getOfficeKey() {
	# 32-bit versions
	$key = Search-RegistryKeyValues "hklm:\software\microsoft\office" "digitalproductid"
	if ($key -eq $null) {
		# 64-bit versions
		$key = Search-RegistryKeyValues "hklm:\software\Wow6432Node\microsoft\office" "digitalproductid"
		if ($key -eq $null) {Write-Host "MS Office is not installed.";break}
	}
	
	$valueData = (Get-ItemProperty $key).digitalproductid[52..66]
	
	# decrypt base24 encoded binary data 
	$oproductKey = ""
	$chars = "BCDFGHJKMPQRTVWXY2346789"
	for ($i = 24; $i -ge 0; $i--) { 
		$r = 0 
		for ($j = 14; $j -ge 0; $j--) { 
			$r = ($r * 256) -bxor $valueData[$j] 
			$valueData[$j] = [math]::Truncate($r / 24)
			$r = $r % 24 
		} 
		$oproductKey = $chars[$r] + $oproductKey 
		if (($i % 5) -eq 0 -and $i -ne 0) { 
			$oproductKey = "-" + $oproductKey 
		} 
	}
	Return $oproductKey
}

function getOfficeID($Arch) {
	if ($Arch -Like '*32*') {
		$Path = "hklm:\software\microsoft\office"
	}
	elseif ($Arch -Like '*64*') {
		$Path = "hklm:\software\Wow6432Node\microsoft\office"
	}
	Get-ChildItem $Path -recurse -ea SilentlyContinue | Where-Object {(Get-ItemProperty -Path $_.PsPath -ea SilentlyContinue) -match "digitalproductid" -eq $True } |  ForEach-Object {$reg = $_.PsPath}
	$Office = (Get-ItemProperty $reg).productid
	Return $Office
}

$Key = getOSVersion
$OfficeKey = getOfficeKey
$OfficeID = getOfficeID($Arch)

addPC
varPCTasks