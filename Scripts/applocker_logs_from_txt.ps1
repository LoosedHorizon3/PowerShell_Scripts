$path = Read-Host "Enter path location on text file (but not the text file itself)"
$computers = Get-Content "$path\computers.txt".Trim()
$computerArray = @()

foreach($device in $computers){
    Write-Host -ForegroundColor Green "Reading text file, please wait..."
    $computerArray += $device
}

foreach($devices in $computerArray){
    if(Test-NetConnection $devices -ErrorAction SilentlyContinue){
        Write-Host "$devices is online. Getting Applocker events for this device"
        Write-Host "Please wait..."
        Get-WinEvent -LogName Microsoft-Windows-AppLocker* -ComputerName $devices | Select-Object Message, Id, LogName, MachineName, UserId, TimeCreated | Where-Object {$_.id -ne 8001} |
        Export-Csv "$path\applocker_results.csv" -Append -NoTypeInformation
    }
    else{
        Write-Host "$devices is offline. Cannot get AppLocker events on this device" -ForegroundColor Red
    }
}
