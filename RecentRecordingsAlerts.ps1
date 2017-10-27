#Author: Alex Dobrovansky
#Last updated: 27 Oct 17

#When run, this script sends a message to slack telling me the number of .mp3 files larger than 50KB, smaller than 50KB, and 1 or 0KB that are less than 2 hours old.
#I use this to know if .mp3 files are being generated in the specified location.
#I orginally had it searching the entire directory, but because the folder structure has the day, month, and year in it, I am able to restrict it for faster searching
#It then sends two messages: one in a very human readable format to my alerts channel, and another to my slack bot in a csv format
#I have then setup the slack bot to monitor these messages and alert me if anything is unusual

#CHANGE THESE VARIABLES!!!
$site = "Site 1" #Name of site (becomes the Slack username)
$date = (Get-Date).AddHours(-2) #How recent to search for
$size = 50KB #Sorting size
$locations = "C:\Recordings", "C:\Users" #Location of recordings (
$fileType = ".mp3" #Filetype of recordings




#create start of message
$body = "Looking for files created after $($date.ToUniversalTime()) UTC" + "`n `n"

#checks the number of recordings in each location
foreach ($loc in $locations){
    $botBody = "" #reset botBody

    $botBody = $site + "," + $loc + "," # add the site and location before it gets changed
    
    #limit the location that it searchs to current day
    $loc += "$(get-date -UFormat '\%Y\%m\%d\')"

    #search for files
    $objects = Get-ChildItem -Path $loc -Recurse | Where-Object {$_.Extension -eq $fileType -and $_.LastWriteTime -gt $date}
    $largeObjects = ($objects | Where-Object {$_.Length -gt $size} | Measure-Object).count
    $smallObjects = ($objects | Where-Object {$_.Length -lt $size} | Measure-Object).count
    $emptyObjects = ($objects | Where-Object {$_.Length -eq 0KB} | Measure-Object).count

    #send msg in csv to bot
    $botBody += "$largeObjects,$smallObjects,$emptyObjects,$date"
    Send-SlackMsg -Username $site -Body $botBody -webhookURL "https://hooks.slack.com/services/" #PS bot

    #append msg to body
    $body += "Recordings in $loc"`
        + "`n"`
        + "Recordings larger than $($size/1KB) KB: $largeObjects"`
        + "`n"`
        + "Recordings smaller than $($size/1KB) KB: $smallObjects"`
        + "`n"`
        + "Recordsings size 0KB or 1KB: $emptyObjects"`
        + "`n `n"`
}

#check free space
$drive = $locations[0][0]
$used,$free = Get-PSDrive $drive | ForEach-Object { $_.Used, $_.Free}
$percent = ($used/($used+$free)).ToString("P")
$body += "Percent of drive $drive used: $percent"

#send function
function Send-SlackMsg {
    param($Username,$Body, $webhookURL = "https://hooks.slack.com/services/") #alerts channel. insert your webhook here 

    $message = @{
        username = $Username
        text = $Body
    }
    $json = $message | ConvertTo-Json
    Invoke-RestMethod -Method Post -Uri $webhookURL -Body $json
}

#send message
Send-SlackMsg -Username $site -Body $body
