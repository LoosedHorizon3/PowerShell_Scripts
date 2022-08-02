# This script will generate all the users within an Active Directory group and can check:
#     - When their passowrd expires
#     - When their passowrd was last updated
#     - When their account expires
# and save it to a csv with the group and also outputs it the screen
# Author: Nicholas Stevenson
# Created: ‎Thursday, ‎19 ‎May ‎2022

##################################################################################################################################################################

# prompting the user for an AD group the check
$group = Read-Host -Prompt "Enter AD Group name "
# Defining where the csv file will be saved
$OutputFile = "F:\Scripts\Outputs\$group.csv"

#This will ask if you wnat to check the users password details (i.e. last set, expried)
$PasswordandAccountExpire = Read-Host "Do you want to check when password and account expires (Type Y to check or enter to contuine)"

# This part of the script shows the password details
if($PasswordandAccountExpire -eq "y"){
    Get-ADGroupMember -Identity $group | ForEach-Object {
        Get-ADUser -Identity $_.SamAccountName -Properties Name, GivenName, Surname, DisplayName, PasswordNeverExpires, PasswordLastSet, AccountExpirationDate | 
        Select-Object -Property DisplayName, Name, GivenName, Surname,PasswordNeverExpires, PasswordLastSet, AccountExpirationDate
        } | Export-Csv $OutputFile -NoTypeInformation
    }

# This part of the script will not shows the password details
else{
    Get-ADGroupMember -Identity $group | ForEach-Object {
        Get-ADUser -Identity $_.SamAccountName -Properties Name, GivenName, Surname, DisplayName | 
        Select-Object -Property DisplayName, Name, GivenName, Surname
        } | Export-Csv $OutputFile -NoTypeInformation
    }

# This is importing the newly created csv file so we can output it to the screen
Import-Csv -Path "F:\Scripts\Outputs\$group.csv" | Format-Table -AutoSize