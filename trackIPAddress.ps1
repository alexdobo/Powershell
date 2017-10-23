#This script gets the public IP address of the computer to a .csv file
#It is designed to be run by task scheduler

$ip = Invoke-RestMethod http://ipinfo.io/json | Select -exp ip
"$ip, $(Get-Date)" | Out-File IPtrackr.csv -append
