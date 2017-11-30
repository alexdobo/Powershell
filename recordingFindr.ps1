#Alex Dobrovansky
#30 Nov 17

#cd path of file first

#v1
#$pattern = Read-Host -Prompt "Enter a phone number or time"
#Get-ChildItem -recurse | Select-String -pattern $pattern | group path | select name


#v2

#define search attributes:
$minStartTime = "09:15:00" #hh:mm:ss
$minStartTime = [datetime]$minStartTime
$maxStartTime = "10:15:00" #hh:mm:ss
$maxStartTime = [datetime]$maxStartTime
$minDuration = New-TimeSpan -minutes 5
$maxDuration = New-TimeSpan -Minutes 10
$phoneNumber = "3211234567"#3211234567
$direction = "INC" #OUT or INC or ""
$matchList = New-Object System.Collections.ArrayList



#all the logic:
$PV5s = Get-ChildItem -recurse | Where-Object {$_.Extension -eq ".pv5"}
foreach ($file in $PV5s){
    [xml]$xml = [regex]::Match($(get-content $file),"<XML>.+<\/XML>").Value
    $file.Name
    #time
    if($minStartTime -and $maxStartTime){
        if(($minStartTime -le [datetime]$xml.XML.Start) -and ($maxStartTime -ge [datetime]$xml.XML.Start)){
            "match both time"
            $timeBool = $True
        }else{
            "outside range"
            $timebool = $False
        }
    }elseif($minStartTime){
        if($minStartTime -le [datetime]$xml.XML.Start){
            "min start time match"
            $timeBool = $True
        }else{
            "too early"
            $timeBool = $False
        }
    }elseif($maxStartTime){
        if($maxStartTime -ge [datetime]$xml.XML.Start){
            "max start time match"
            $timeBool = $True
        }else{
            "too late"
            $timeBool = $False
        }
    }else{
        $timeBool = $True
    }

    #duration
    $callDuration = [datetime]$xml.XML.End - [datetime]$xml.XML.Start
    if($minDuration -and $maxDuration){
        if(($minDuration -le $callDuration) -and ($maxDuration -ge $callDuration)){
            "within duration"
            $durationBool = $True
        }else{
            "out of duration"
            $durationBool = $False
        }
    }elseif($minDuration){
        if($minDuration -le $callDuration){
            "min duration"
            $durationBool = $True
        }else{
            "too short"
            $durationBool = $False
        }
    }elseif($maxDuration){
        if ($maxDuration -ge $callDuration) {
            "max duration"
            $durationBool = $True
        }else{
            "too long"
            $durationBool = $False
        }
    }else{
        $durationBool = $True
    }
    
    #phone number
    if($phoneNumber){
        if ($xml.XML.CLI -like $phoneNumber) {
            "CLI Match"
            $phoneNumberBool = $True
        }elseif($xml.XML.DDI -like $phoneNumber){
            "DDI match"
            $phoneNumberBool = $True
        }else{
            "number didn't match"
            $phoneNumberBool = $False
        }
    }else{
        $phoneNumberBool = $True
    }

    #direction
    if ($direction) {
        if ($direction -eq $xml.XML.DIRECTION) {
            "Direction"
            $directionBool = $True
        }else{
            "wrong way go back"
            $directionBool = $False
        }
    }else{
        $directionBool = $True
    }

    if($timeBool -and $durationBool -and $phoneNumberBool -and $directionBool){
        $matchList.Add($file.Name)
        "ITS A MATCH!"
    }
}
$matchList
