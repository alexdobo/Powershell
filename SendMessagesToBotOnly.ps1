#Author: Alex Dobrovansky
#Last updated: 27 Oct 17



#CHANGE THESE VARIABLES!!!
$site = "Site 1" #Name of site (becomes the Slack username)
$date = (Get-Date).AddHours(-2) #How recent to search for
$size = 50KB #Sorting size
$recLocations = "C:\Recordings\loc1", "C:\Recordings\loc2" #Location of recordings ("loc1","loc2")
$colRecLocations = "C:\Call data\loc1","C:\Call Data\loc2" #Location of .col and .rec files.     MUST BE IN SAME ORDER AS ABOVE
$fileType = ".pv5" #Filetype of recordings



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





#get date from col and rec files
$lastModCol = New-Object System.Collections.ArrayList
$lastModRec = New-Object System.Collections.ArrayList

#send all of the information to the bot
#send number of files and last col date s

if ($recLocations.Count -ne $colRecLocations.Count){ #must be same number of reclocs and colRecLocs
    Throw "The number of Recording locations and .col file location does not match!"
}else{
    $botBody=""
    for ($i=0; $i -lt $recLocations.Count; $i++){
        
        #Site,Loc,Large,Small,Empty,Date (UTC),Free Space,Last Col Date,Last Rec Date
        $botBody += "$site,$($recLocations[$i]),"

        $loc = $recLocations[$i] + "$(get-date -UFormat '\%Y\%m\%d\')"
    
        #search for files
        $objects = Get-ChildItem -Path $loc -Recurse | Where-Object {$_.Extension -eq $fileType -and $_.LastWriteTime -gt $date}
        $largeObjects = ($objects | Where-Object {$_.Length -gt $size} | Measure-Object).count
        $smallObjects = ($objects | Where-Object {$_.Length -lt $size} | Measure-Object).count
        $emptyObjects = ($objects | Where-Object {$_.Length -eq 0KB} | Measure-Object).count

        #create botmsg
        $botBody += "$largeObjects,$smallObjects,$emptyObjects,$($date.ToUniversalTime()),"

        $drive = $loc[0]
        $used,$free = Get-PSDrive $drive | ForEach-Object { $_.Used, $_.Free}
        $percent = ($used/($used+$free)).ToString("P")
        $bodyBot += "$percent,"

        $bodyBot += "$((Get-ChildItem -path $colRecLocations[$i] -Recurse | where {$_.Extension -eq ".col"} | sort LastWriteTime -Descending | select -First 1).LastWriteTime.ToUniversalTime()),"
        $bodyBot += "$((Get-ChildItem -path $colRecLocations[$i] -Recurse | where {$_.Extension -eq ".rec"} | sort LastWriteTime -Descending | select -First 1).LastWriteTime.ToUniversalTime()) `n"



        
    }
    #sends all the data about site in one message
    Send-SlackMsg -Username $site -Body $botBody -webhookURL "https://hooks.slack.com/services/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" # @PS bot
}

