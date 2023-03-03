# Author: Nicholas Stevenson
# Date: 13 Feburary 2023
# This script will ask for list of devices and then check their event viewer logs for AppLocker Events
##################################################################################################################################

# Defining the variable computers as a string with an empty character
$computers = " "
# Defining an empty array to be filled in later
$computerArray = @()

# WHILE string does not equal nothing ask for input ELSE move on
while($computers -ne ""){
    $computers = Read-Host "Enter PC name (Press enter on a new line to exit)"
    # Add string to array
    $computerArray += $computers
}
# Prompt for CSV path location to be stored in
$path = Read-Host "Enter path name to store the csv file on the local device"

# FOR each computer in array get AppLocker event logs
foreach($device in $computerArray){
    Write-Host "$device Applocker events are:"
    # Get all AppLocker blocking events and export all logs to a CSV file at the end
    Get-WinEvent -LogName Microsoft-Windows-AppLocker* -ComputerName $device | Select-Object Message, Id, LogName, MachineName, UserId, TimeCreated |
    Where-Object {$_.id -eq 8003 -or $_.id -eq 8006 -or $_.id -eq 8004 -or $_.id -eq 8007} | Export-Csv "$path\applocker_blocked_results.csv" -Append -NoTypeInformation
    # Reading the CSV file out onot the screen
    Get-Content $path\applocker_blocked_results.csv
}
