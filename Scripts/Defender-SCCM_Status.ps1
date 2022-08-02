# This script will prompt the user to enter in the name of a computer one at a time and form a list of computers.
# The script will then check the status and version of the SCCM Client and the Microsoft 365 Defender Client and output the result.
# This script will also use the local machine as a referance PC to check the version of SCCM Client and the Microsoft 365 Defender Client for a baseline
# Author: Nicholas Stevenson
# Created: Tuesday, ‎26 ‎July ‎2022

##################################################################################################################################################################

$computers = " "
$computerArray = @()
$array = @()
$refComputer = hostname
$referencePCDefender = Get-MpComputerStatus | Select AMEngineVersion, AMRunningMode, AntispywareSignatureLastUpdated -ErrorAction Continue
$referencePCSCCM = (Get-WMIObject -ComputerName $refComputer -Namespace root\ccm -Class SMS_Client).ClientVersion

# This will create the list of computers entered
while($computers -ne ""){
    $computers = Read-Host "Enter PC name (Press enter on a new line to exit)"
    $computerArray += $computers
}

# This will use the the PC used to run the script to use as a referance
if($referencePCSCCM -ne $null){
            $Object = New-Object psobject -Property (@{
                "Computer Name" = $refComputer
                "SSCM Client Version" = $referencePCSCCM
                "Defender Client Status/Version" = $referencePCDefender
            })
	    # Adding the output of the referance computer to the list that will hold the final output
            $array += $Object
        }


# This will lop through the list of computers and check if the computer is accessible over the network
# and then check the SCCM and Defender status
foreach($computer in $computerArray){
    try{
        # Checking to see if the device is accessible over the network
        if($computer -ne ""){
            # If the device is not accessiable over the network output the below
            if((Test-Path \\$computer\c$) -match "False"){
                Write-Warning "Failed to connect to $computer"
            }
            Else
            {
                # Setting the values of the variables to null so we can use it later in the script
	            $SCCMClient = $null
                $Object = $null
                $DefenderStatus = $null
        	
            	# Defining how to retrieve the status of MS Defender and SCCM
	            $DefenderStatus = Get-MpComputerStatus -CimSession $computer | Select AMEngineVersion, AMRunningMode, AntispywareSignatureLastUpdated -ErrorAction Continue
                $SCCMClient = (Get-WMIObject -ComputerName $computer -Namespace root\ccm -Class SMS_Client).ClientVersion
	
	
                if($SCCMClient -ne $null){
                    $Object = New-Object psobject -Property (@{
                        "Computer Name" = $computer
                        "SSCM Client Version" = $SCCMClient
                        "Defender Client Status/Version" = $DefenderStatus
                    })
    	        # Adding the outputs of the above to the list that hold the final output
                    $array += $Object
                }
                else
                {
                    $Object = New-Object psobject -Property (@{
                        "Computer Name" = $computer
                        "SSCM Client Version" = "(Null)"
                        "Defender Client Status" = "(Null)"
                 })
                          
                    $array += $Object
                }
            }

        }
    }
    catch [Microsoft.PowerShell.Cmdletization.Cim.CimJobException]{
        Write-Warning "WinRM is not enabled on $computer"
    }
}


# Outputting the final list to screen
if($array){
    Write-Host "The top computer name will be the local PC the script is run on to be used as a referance" -ForegroundColor White -BackgroundColor Black
    Write-Output $array | Format-Table -AutoSize
}