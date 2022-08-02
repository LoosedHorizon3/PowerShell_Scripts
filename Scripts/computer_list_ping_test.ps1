# This script will take a text file that contains list of computer names, ping them, and output weather the PC is turned on and accessiable or turned off
# Author: Nicholas Stevenson
# Created: ‎‎Wednesday, ‎20 ‎July ‎2022

##################################################################################################################################################################

$computers = Get-Content "Copmuter_List.txt"
$computerList = @()

foreach($device in $computers){
    $computerList += $device
}
 
foreach($computer in $computerList)
     {
     if (Test-Connection  $x -Count 1 -ErrorAction SilentlyContinue){
         Write-Host "$computer is up"
         }
     else{
         Write-Host "$computer is down"
         }
     }
     