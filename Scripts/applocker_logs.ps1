# Author: Nicholas Stevenson
# Date: 13 Feburary 2023
# This script gets AppLocker events from a list of computers, filters them, saves them to a CSV file, and displays the CSV file's contents.
#####################################################################################################################################################

# Initialize variables to be used later in the script
$computers = " "
$computerArray = @()

# Output script purpose to console
Write-Host "This script gets AppLocker events from a list of computers, filters them, saves them to a CSV file, and displays the CSV file's contents." -ForegroundColor Green

# Prompt the user to enter computer names to search for AppLocker events, and add them to the $computerArray
while($computers -ne ""){
    $computers = Read-Host "Enter PC name (Press enter on a new line to exit)"
    $computerArray += $computers
}

# Prompt the user to enter the path to store the CSV file on the local device
$path = Read-Host "Enter path name to store the csv file on the local device"

# Loop through the $computerArray and search for AppLocker events on each computer
foreach($device in $computerArray){
    # Check if the computer is online
    if(Test-NetConnection $device -ErrorAction SilentlyContinue){
        Write-Host "$device is online. Getting Applocker events for this device"
        Write-Host "Please wait..."
        # Use Get-WinEvent cmdlet to search for AppLocker events on the remote computer, and select relevant properties
        Get-WinEvent -LogName Microsoft-Windows-AppLocker* -ComputerName $device | Select-Object Message, Id, LogName, MachineName, UserId, TimeCreated |
        # Exclude event ID 8001 as it's an informational event and not useful for this search
        Where-Object {$_.id -ne 8001} |
        # Export the selected properties for each AppLocker event to a CSV file at the specified path, and append each result to the existing file
        Export-Csv "$path\applocker_results.csv" -Append -NoTypeInformation
    }
    else{
        # Output a message if the computer is offline
        Write-Host "$device is offline. Cannot get AppLocker events on this device" -ForegroundColor Red
    }
}
