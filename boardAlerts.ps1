#Author: Alex Dobrovansky
#Date: 08 Nov 17
#This script reads a .rec file which is formatted as a .csv
#it does some math, then sends an alert

$path = "C:\CallData\"
$divider = 30
$site = "Test site 1"



function Send-SlackMsg {
    param($Username,$Body,$webhookURL = "https://hooks.slack.com/services/"<# #Alerts Channel #>)

    $message = @{
        username = $Username
        text = $Body
    }
    $json = $message | ConvertTo-Json

    
    Invoke-RestMethod -Method Post -Uri $webhookURL -Body $json
}

function Write-Alert {
    param(
        $Status
    )
    $Private:value = ""
    foreach($j in $status){
        if((-not $j.Above) -and (-not $j.Below)){
            $value += "We are not getting calls on the board $($j.board) on all channels `n"
        }elseif ((-not $j.Above)) {
            $value += "We are not getting calls on the board $($j.board) on channels above $divider `n"
        }elseif ((-not $j.Below)) {
            $value += "We are not getting calls on the board $($j.board) on channels below $divider `n"
        }
    }

    return $value

}
function Check-Status{
    param(
        $Status
    )
    $Private:value = $true
    foreach($i in $status){
        if(-not $i.Above){
            $value = $false
        }
        if (-not $i.Below){
            $value = $false
        }
    }
    return $value
}
function Get-Status{
    param(
        $path,
        $divider
    )
    $PRIVATE:recFiles = ""
    $PRIVATE:csv=""
    $PRIVATE:listStatus = @()
    $PRIVATE:statusObj = new-object psobject -Property (@{"Board"="";"Above"="";"Below"=""})
    $recFiles = Get-ChildItem  -path $path| Where-Object {$_.Extension -eq ".rec"} | Sort-Object -Property LastWriteTime
    $csv = import-csv -path "$path\$($recFiles[-1].Name)" | Select-Object Board,Channel | Group-Object -Property Board
    
    foreach ($board in $csv){
        $statusObj.Below = $false
        $statusObj.Above = $false
        $statusObj.Board = $board.Name
        $PRIVATE:stats = $board.Group.Channel | Measure-Object -Maximum -Minimum
        if ($stats.Minimum -le $divider){
            $statusObj.Below = $true    
        }
        if ($stats.Maximum -gt $divider) {
            $statusObj.Above = $true
        }
        $listStatus += $statusObj.psobject.Copy()
    }
    return $listStatus
}








$stat = Get-Status -path $path -divider $divider
if((Check-Status -status $stat)){
    "Everything is working"
}else {
    "send alert"
    $stat
    $alert = Write-Alert -Status $stat
    $alert
    #Send-SlackMsg -Username $site -Body $alert
}

