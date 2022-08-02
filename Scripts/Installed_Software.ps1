# This script will prompt the user to enter in the name of a computer and then query the regestry to check for installed software
# Author: Nicholas Stevenson
# Created: ‎Friday, ‎24 ‎June ‎2022

##################################################################################################################################################################

# This will prompt the user for the user to enter the name of the computer
$computer = Read-Host "Enter computer name"

# This will qurery the regestry of the device to and output the software installed on the PCs
Write-Host "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
Invoke-Command -Computer $computer {Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
    Select DisplayName, Publisher, InstallDate, DisplayVersion | Format-Table -AutoSize}

Write-Host "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
Invoke-Command -Computer $computer {Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | 
    Select DisplayName, Publisher, InstallDate, DisplayVersion | Format-Table -AutoSize}