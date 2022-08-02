# This script will take the users input and attempt to enable powershell remoting
# Requirements: PsExec https://docs.microsoft.com/en-us/sysinternals/downloads/psexec
# Author: Nicholas Stevenson
# Created: Monday, ‎25 ‎July ‎2022

##################################################################################################################################################################

# Getting the input from the user and assigning it to a variable
$device = Read-Host "Enter computer name"

# This will use PSExec on the device to enable powershell remoting on the device specified
C:\PSTools\PsExec.exe \\$device -h -s powershell.exe Enable-PSRemoting -Force