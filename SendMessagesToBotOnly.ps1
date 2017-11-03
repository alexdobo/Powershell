#Author: Alex Dobrovansky
#Last updated: 03 Nov 17



#CHANGE THESE VARIABLES!!!
$site = "Site 1" #Name of site (becomes the Slack username)
$date = (Get-Date).AddHours(-2) #How recent to search for
$size = 50KB #Sorting size
$recLocations = "C:\Documents\WindowsPowerShell\Slack Bot"#, "C:\Recordings\loc2" #Location of recordings ("loc1","loc2")
$colRecLocations = "C:\WindowsPowerShell\Slack Bot"#,"C:\Call Data\loc2" #Location of .col and .rec files.     MUST BE IN SAME ORDER AS ABOVE
$fileType = ".mp3"


#send function
function Send-SlackMsg {
    param($Username,$Body,$webhookURL = "https://hooks.slack.com/services/"<# #Alerts Channel #>)

    $message = @{
        username = $Username
        text = $Body
    }
    $json = $message | ConvertTo-Json

    
    Invoke-RestMethod -Method Post -Uri $webhookURL -Body $json
}



#send all of the information to the bot
#send number of files and last col date s

if ($recLocations.Count -ne $colRecLocations.Count){ #must be same number of reclocs and colRecLocs
    Throw "The number of Recording and .col file locations does not match!"
}else{
    #Site,Loc,Large,Small,Empty,Date (UNIX),Free Space,Last Col Date,Last Rec Date
    
    if ($recLocations.Count -eq 1){
        $botBody=""
        $botBody += "$site,$($recLocations),"

        $loc = $recLocations + "$(get-date -UFormat '\%Y\%m\%d\')"
        
        #search for files
        $objects = Get-ChildItem -Path $loc -Recurse | Where-Object {$_.Extension -eq $fileType -and $_.LastWriteTime -gt $date}
        $largeObjects = ($objects | Where-Object {$_.Length -gt $size} | Measure-Object).count
        $smallObjects = ($objects | Where-Object {$_.Length -lt $size} | Measure-Object).count
        $emptyObjects = ($objects | Where-Object {$_.Length -eq 0KB} | Measure-Object).count

        #create botmsg
        $botBody += "$largeObjects,$smallObjects,$emptyObjects,$([int][double]::Parse((Get-Date ($date).touniversaltime() -UFormat %s)))," #convert to UNIX time

        $drive = $loc[0]
        $used,$free = Get-PSDrive $drive | ForEach-Object { $_.Used, $_.Free}
        $percent = ($used/($used+$free))
        $botBody += "$percent,"


        $lastColTime = (Get-ChildItem -path $colRecLocations -Recurse | where {$_.Extension -eq ".col"} | sort LastWriteTime -Descending | select -First 1).LastWriteTime
        $lastRecTime = (Get-ChildItem -path $colRecLocations -Recurse | where {$_.Extension -eq ".rec"} | sort LastWriteTime -Descending | select -First 1).LastWriteTime

        $botBody += "$([int][double]::Parse((Get-Date ($lastColTime).ToUniversalTime() -UFormat %s))),"
        $botBody += "$([int][double]::Parse((Get-Date ($lastRecTime).ToUniversalTime() -UFormat %s)))"
        $botBody
        #sends all the data about site in one message
        Send-SlackMsg -Username $site -Body $botBody -webhookURL "https://hooks.slack.com/services/" # @PS bot

    }else {
        for ($i=0; $i -lt $recLocations.Count; $i++){
            $botBody = ""
        
            $botBody += "$site,$($recLocations[$i]),"

            $loc = $recLocations[$i] + "$(get-date -UFormat '\%Y\%m\%d\')"
    
            #search for files
            $objects = Get-ChildItem -Path $loc -Recurse | Where-Object {$_.Extension -eq $fileType -and $_.LastWriteTime -gt $date}
            $largeObjects = ($objects | Where-Object {$_.Length -gt $size} | Measure-Object).count
            $smallObjects = ($objects | Where-Object {$_.Length -lt $size} | Measure-Object).count
            $emptyObjects = ($objects | Where-Object {$_.Length -eq 0KB} | Measure-Object).count

            #create botmsg
            $botBody += "$largeObjects,$smallObjects,$emptyObjects,$([int][double]::Parse((Get-Date ($date).touniversaltime() -UFormat %s)))," #convert to UNIX time

            $drive = $loc[0]
            $used,$free = Get-PSDrive $drive | ForEach-Object { $_.Used, $_.Free}
            $percent = ($used/($used+$free))
            $botBody += "$percent,"


            $lastColTime = (Get-ChildItem -path $colRecLocations[$i] -Recurse | where {$_.Extension -eq ".col"} | sort LastWriteTime -Descending | select -First 1).LastWriteTime
            $lastRecTime = (Get-ChildItem -path $colRecLocations[$i] -Recurse | where {$_.Extension -eq ".rec"} | sort LastWriteTime -Descending | select -First 1).LastWriteTime

            $botBody += "$([int][double]::Parse((Get-Date ($lastColTime).ToUniversalTime() -UFormat %s))),"
            $botBody += "$([int][double]::Parse((Get-Date ($lastRecTime).ToUniversalTime() -UFormat %s)))"
            $botBody
            #sends all the data about site in one message
            Send-SlackMsg -Username $site -Body $botBody -webhookURL "https://hooks.slack.com/services/" # @PS bot
        }

    }

}

