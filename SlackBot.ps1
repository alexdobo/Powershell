#Author: Alex Dobrovansky
#Date: 03 Nov 17


$token = "xoxb"
$tolerance = 5
$alertsChannel = "CK"
$monitorList = New-Object System.Collections.ArrayList
import-csv "monitorList.csv" -Header "Name" | foreach{$monitorList.add($_.Name)}
$siteList = New-Object System.Collections.ArrayList
import-csv "recordingData.csv" | Group-Object Site | foreach{$siteList.add($_.Name)}

function Send-SlackMsg {
    param($Channel,$Body,$attachments)

    $message = @{
        id = "ps-bot"
        type = "message"
        text = $Body
        channel = $channel
    }
    $json = $message | ConvertTo-Json

    $array = @()
    $encoding=
    $array = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Msg = New-Object System.ArraySegment[byte]  -ArgumentList @(,$Array)

    if(!$attachments){
        $Conn = $WS.SendAsync($Msg, [System.Net.WebSockets.WebSocketMessageType]::Text, [System.Boolean]::TrueString, $CT)
    }else{
        $message.as_user = "false"
        $message.username = "ps-bot"
        $message.token = "xoxb-"
        $message.attachments = $attachments
        Invoke-WebRequest -Method post -Uri "https://slack.com/api/chat.postMessage" -body $message | out-file -filepath "slack.log" -append
    }

    


}


function Within-Tolerances {
    param ($tol, $value, $avg)
    if ($value -ge ($avg - $tolerance) -and $value -le  ($avg + $tolerance))    {
        return $True
    }else {
        return $False
    }
}

#Web API call starts the session and gets a websocket URL to use.
$RTMSession = Invoke-RestMethod -Uri https://slack.com/api/rtm.start -Body @{token="$Token"}

"I am $($RTMSession.self.name)"
$rtmCount = 0
$msgCount = 0
Try{

    Do{
        $WS = New-Object System.Net.WebSockets.ClientWebSocket                                                
        $CT = New-Object System.Threading.CancellationToken
        
        #start webstocket connection
        $Conn = $WS.ConnectAsync($RTMSession.URL, $CT)                                                  
        
        #wait until connection is made
        While (!$Conn.IsCompleted) { Start-Sleep -Milliseconds 100 }
        
        "Connected to $($RTMSession.URL)"

        #creates an array to store received files
        $size = 1024
        $Array = [byte[]] @(,0) * $Size
        $Recv = New-Object System.ArraySegment[byte] -ArgumentList @(,$Array)

        while ($ws.State -eq "Open"){

            if((get-date).TimeOfDay.TotalSeconds -ge 43200 -and (get-date).TimeOfDay.TotalSeconds -lt 43220 -and (get-date).DayOfWeek -ne "Saturday" -and (get-date).DayOfWeek -ne "Sunday"){
                $dateCsv = ""
                $dateCsv = Import-Csv recordingData.csv

                foreach ($site in $monitorList){
                    $date = ($dateCsv | Where-Object {$_.Site -eq $site}).Date[-1] #gets the latest date
                    if ($date -lt ((get-date).addHours(-25))) { #checks within last 25 hours
                         send-SlackMsg -body "I haven't heard from $site since $date" -channel $alertsChannel
                    }
                    $date = ""
                }
                sleep 21
            }

            #deal with data
            if ($rtmCount -gt 5 -and $msgCount -gt 0){ #waits for 10 cycles before sending any messages
                #because every line of data is a new message, this joins them all so we don't get spammed
                $rtmCount = 0
				$msgCount = 0
				$listData

                $csv = ""
                $csv = Import-Csv "recordingData.csv"
                $bigMsg = "For the site "+ $listData[0].site
                $attach = "["
                foreach ($data in $listData){
                    #for each line of data
                    $collec = $csv | Group-Object -property Site,Location | Where-Object {$_.Name -eq "$($Data.site), $($Data.location)"}
                    
                    $check1 = Within-Tolerances -tol $tolerance -avg $($collec.Group.Large | measure -Average).Average -value $Data.large
                    $check2 = Within-Tolerances -tol $tolerance -avg $($collec.Group.Small | measure -Average).Average -value $Data.small
                    $check3 = Within-Tolerances -tol $tolerance -avg $($collec.Group.Empty | measure -Average).Average -value $Data.empty

					
                    if (!$check1 -or !$check2 -or !$check3){
                        "somethings's not right"
                        $color = "warning"   
                    }else {
                        $color = "good"
                    }
										
		    if ($data.lastRec -lt (((get-date).AddHours(-2)).toUniversalTime)){ #something about last col file being older than 2 hours
		    	$color = "danger"
		    }
                
                    #Send-SlackMsg -Channel $alertsChannel -Body $bigMsg -attachments $attach

                    $attach += '{
                        "fallback":"Alert for location '+$data.location+'",
                        "color": "'+$color+'",
                        "title": "'+$site.location+'",
                        "fields": [
                            {
                                "title": "Last write to .col file",
                                "value": "<!date^'+ ($data.LastCol) +'^{time} on {date}|you need to update slack>",
                                "short": true
                            },
                            {
                                "title": "Last write to .rec file",
                                "value": "<!date^'+ ($data.LastRec) +'^{time} on {date}|you need to update slack>",
                                "short": true
                            },
                            {
                                "title": "Large Rec",
                                "value": "'+$data.Large+'",
                                "short": true
                            },
                            {
                                "title": "Small Rec",
                                "value": "'+$data.small+'",
                                "short": true
                            },
                            {
                                "title": "Empty Rec",
                                "value": "'+$data.empty+'",
                                "short": true
                            },
                            {
                                "title": "Free Space",
                                "value": "'+([double]($data.FreeSpace)).ToString("P")+'",
                                "short": true
                            }

                        ]

                        },'

                }#end loop

                $attach = $attach.Substring(0,$attach.Length-1) #removes last , from attach
                $attach += "]"
				$attach = $attach -replace "\\", "\\"
                $attach = $attach -replace "\/", "\/"
                $attach
		send-slackmsg -body $bigMsg -attachments $attach -channel $alertsChannel #send the large message
				
                if (!$siteList.contains($Data.site)){
                    $siteList.add($data.site)
                }
            }#end deal with data

            $RTM = ""

            Do {
                $Conn = $WS.ReceiveAsync($recv, $CT)
                while (!$Conn.IsCompleted){ Sleep -Milliseconds 100 }
                $recv.Array[0..($Conn.Result.Count -1)] | ForEach-Object { $RTM = $RTM + [char]$_ }

            } Until ($conn.result.count -lt $size)
            $RTM
            $rtmCount
            if ($RTM){
                if ($msgCount -gt 0) {
					$rtmCount += 1
				}
                $RTM = ($RTM | convertfrom-json)
                Switch ($RTM){
                    {($_.type -eq "message") -and (!$_.reply_to)} {
                        
                        if (($_.text -match "<@$($RTMSession.self.id)>") -or $_.channel.StartsWith("D")) { 
                            #msg sent to the bot
                            #insert response
                            $_.user
                            $matches = ""
                            switch ($_.text){
                                #help
                                {$_ -match ".*help.*"}{
                                    $msg = "*Help* file `n The following is a list of commands that are accepted. The key words are in bold. To call a command, the message must have the key words in it. `n `n"
                                    $msg += "*List sites* `n"
                                    $msg += "This will list all of the sites that have sent data to the me. `n"
                                    $msg += " `n"
                                    $msg += "*List monitor* or *List monitored sites* `n"
                                    $msg += "This will list all of the sites that are being monitored `n"
                                    $msg += "If a site that is being monitored hasn't sent me data within the last 25 hours, I will send a message to the <#C7MJWNARK|alerts> channel `n"
                                    $msg += " `n"
                                    $msg += "*Start monitor* ing 'Site 1' `n "
                                    $msg += "This will add 'Site 1' to the monitor list `n"
                                    $msg += "Do not include the quotation marks `n"
                                    $msg += " `n"
                                    $msg += "*Stop monitor* ing 'Site 1' `n"
                                    $msg += "This will remove 'Site 1' to the monitor list `n"
                                    $msg += "Do not include the quotation marks `n"
                                    $msg += " `n"
                                    $msg += "*Set tolerance* 4 `n"
                                    $msg += "This will set the recording alerts tolerance to 4 `n"
                                    $msg += " `n"
                                    $msg += "Tell me a *joke* `n"
                                    $msg += "This will tell you a Chuck Norris joke `n"
                                    $msg += " `n"
                                    $msg += "Give me an *excuse* `n"
                                    $msg += "This will give you a randmon excuse `n"
                                    $msg += " `n"
                                    $msg += "Send me a *photo* of a *dog* `n"
                                    $msg += "This will send you a dog photo `n"
                                    $msg += " `n"
                                    $msg += "Tell me the *time* `n"
                                    $msg += "This will tell you the time `n"
                                    $msg += " `n"
                                    $msg += "*Help* `n"
                                    $msg += "This will send you this message `n"
                                    send-SlackMsg -body $msg -channel $rtm.channel
                                }
                                #joke
                                {$_ -match ".*joke.*"}{
                                    $joke = ""
                                    $joke = ((Invoke-RestMethod -Method Get -Uri "http://api.icndb.com/jokes/random").value).joke
                                    send-SlackMsg -body $joke -channel $rtm.channel
                                }
                                ##cat fact - not working
                                #{$_ -match ".*cat fact.*"}{
                                #    $catFact = ""
                                #    $catFact = ((Invoke-RestMethod -Method Get -Uri "https://catfact.ninja/fact").value).fact
                                #    send-SlackMsg -body $catFace -channel $rtm.channel
                                #}
                                #excuse
                                {$_ -match ".*excuse.*"}{
                                    $excuse = ""
                                    $excuse = (Invoke-WebRequest http://pages.cs.wisc.edu/~ballard/bofh/excuses -OutVariable excuses).content.split([Environment]::NewLine)[(get-random $excuses.content.split([Environment]::NewLine).count)]
                                    send-SlackMsg -body $excuse -channel $rtm.channel
                                }
                                #dog photo
                                {$_ -match ".*dog.*" -and $_ -match ".*photo.*" -and !($_ -match "Here is a photo of a dog:")}{
                                    $image = (((Invoke-WebRequest -Method get -Uri "http://api.thedogapi.co.uk/v2/dog.php").content) | convertFrom-json).data.url
                                    $attach = "[{'fallback':'Dog Photo','image_url':'$image'}]"
                                    send-slackmsg -body "Here is a photo of a dog:" -attachments $attach -channel $rtm.channel
                                }
                            
                            
                                #tell me about site x
                                #read .col and .rec files
                            
                            
                            
                                {$_ -match ".*time.*"}{
                                    send-SlackMsg -body "The time is $(get-date)" -channel $rtm.channel
                                }
                                {$_ -match "list site.*" -or $_ -match "site list"}{
                                    $body = "Sites: `n"
                                    foreach($x in $siteList){$body += "$x `n"}
                                    send-SlackMsg -body $body -channel $rtm.channel
                                    $body = ""
                                }
                                {$_ -match "list monitored sites" -or $_ -match "list monitor" -or $_ -match "monitor list"}{
                                    $body = "Sites being monitored: `n"
                                    foreach($x in $monitorList){$body += "$x `n"}
                                    send-SlackMsg -body $body -channel $rtm.channel
                                    $body = ""
                                }
                                {$_ -match "start monitor\w{0,3} (.+)"}{
                                    $_ -match "start monitor\w{0,3} (.+)"
                                    if ($monitorList.contains($matches[1])){ #need to write this to a file
                                        send-SlackMsg -body "$($matches[1]) is already on the monitor list" -channel $rtm.channel
                                    }else{
                                        if ($siteList.contains($matches[1])) {
                                            send-SlackMsg -body "The site '$($matches[1])' has been added to the monitor list" -channel $rtm.channel
                                            $monitorList.add($matches[1])
                                            $monitorList | Out-File "monitorList.csv"
                                        }else{
                                        send-SlackMsg -body "The site '$($matches[1])' does not exist." -channel $rtm.channel
                                        }
                                    }
                                }
                                {$_ -match "stop monitor\w{0,3} (.+)"}{
                                    $_ -match "stop monitor\w{0,3} (.+)"
                                    if ($monitorList.contains($matches[1])){
                                        $monitorList.remove($matches[1])
                                        send-SlackMsg -body "The site '$($matches[1])' has been removed to the monitor list" -channel $rtm.channel
                                        $monitorList | Out-File "monitorList.csv"
                                    }else{
                                        send-SlackMsg -body "The monitor list does not contain $($matches[1])" -channel $rtm.channel
                                    }
                                }
                                {$_ -match "set tolerance to (\d+)"}{
                                    $tolerance = $matches[1]
                                    send-SlackMsg -body "The tolerance is now set to $tolerance" -channel $rtm.channel
                                }


                                {$_ -match "(.+),(.+),(\d+),(\d+),(\d+),(\d+),(\d\.\d+),(\d+),(\d+)"}{
                                    #Site,Loc,Large,Small,Empty,Date (UNIX),Free Space,Last Col Date,Last Rec Date
                                    #is a csv, do analysis
                                    $msgCount+=1

                                    $listData = New-Object System.Collections.ArrayList


                                    $splitData = $_ -split ","
                                    $info = @{
                                        "Site" = $splitdata[0];
                                        "Location" = $splitdata[1];
                                        "Large" = $splitdata[2];
                                        "Small" = $splitdata[3];
                                        "Empty" = $splitdata[4];
                                        "Date" = $splitdata[5];
                                        "FreeSpace" = $splitdata[6];
                                        "LastCol" = $splitdata[7];
                                        "LastRec" = $splitdata[8];
                                    }
                                    $listdata.add($info)
                                    $info
                                    $_ | Out-File "recordingData.csv" -Append

                                }
                                
                                {$_ -match "(.+),(.+),(\d+),(\d+),(\d+),(\d\d?\/\d\d?\/\d{4}\s.+)"}{
                                    #Site,Loc,Large,Small,Empty,Date (UTC),Free Space
                                    #old style. still here for legacy support
                                    #just writes to file                                
                                    $_ | Out-File "recordingData.csv" -Append
                                    #and then complains about it
                                    $splitData = $_ -split ","
                                    send-slackmsg -body "I received some data from a site using an old data format. Please update $($splitData[0])" -channel $alertsChannel #send the large message

                                }
                                default { 
                                    "Msg ignored, no response"
                                    $_ 
                                }

                            }
                        
                        } Else {

                            "Msg ignored, not sent to me"
                            $_.text
                        
                        }

                    }
                    {$_.type -eq "reconnect_url"} {$RTMSession.URL = $RTM.url }

                    default {"No action for $rtm.type"}
                }
            }
        
        
        }
    } Until (!$conn)


}Finally {
    if ($WS) {
        "closing websocket"
        $ws.Dispose()

    }

}
