# This script will get the users name, user id, office, and what access they have in AD
# Author: Nicholas Stevenson

##################################################################################################################################################################

# This will ask the user to input the username of the user if know
# If unknown, hitting enter on a blank line will take you to check the display name
$CheckUsername = Read-Host "Do you know the Username (Y to continue or enter to move to the next stage)"

if($CheckUsername -eq "y"){
    $Username = Read-Host "Enter Username"
    Get-ADUser -f "Name -eq '$Username'" -Properties Office, DisplayName, Title | Select Name, DisplayName, GivenName, Surname, Title, Office | Format-List
 
    # This part will generate the users Active Directory access if "y" is prompted otherwise it will exit the script
    $UserAdGroup = Read-Host "Do you want to check the AD access (Y to generate access or enter to exit)"

    if($UserAdGroup -eq "y"){
        Get-ADPrincipalGroupMembership $Username | Select name
    }
}

Else{
    # If the username is not know, this will prompt for the users Display Name
    $CheckDisplayname = Read-Host "Do you know the Displayname (Y to continue or enter to exit)"
        if($CheckDisplayname -eq "y"){
            $Displayname = Read-Host "Enter Displayname"
            Get-ADUser -f "DisplayName -eq '$DisplayName'" -Properties Office, DisplayName, Title | Select Name, DisplayName, GivenName, Surname, Title, Office | Format-List
            
            # This part will generate the users Active Directory access if "y" is prompted otherwise it will exit the script
            $UserAdGroup = Read-Host "Do you want to check the AD access (Y to generate access or enter to exit)"
           
            # The username is then prompted as thee script uses the name field in Active Directory to pull this data
            if($UserAdGroup -eq "y"){
                $Name = Read-Host "Enter the username"
                Get-ADPrincipalGroupMembership $Name | Select name
            }
        }
}