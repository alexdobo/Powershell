#Alex Dobrovansky
#02 Jan 17
#Trello Automation


Function MsgBox ($Text,$Err = "Error"){
    $null = @(
        Add-Type -AssemblyName System.Windows.Forms
        Write-Host $text
        [System.Windows.Forms.MessageBox]::Show($Text,"Trello Inuputr","Ok",$Err)
    )
}


Function CheckForTrellOps(){
    if (get-module -ListAvailable -Name TrellOps){
        return $True
    }else{
        try{
            Write-Host "Installing TrellOps"
            Write-Host "You may be prompted for confirmation. Select Yes"
            install-module "TrellOps" -Scope CurrentUser -Force
        }catch{
            msgBox "Failed to install TrellOps"
            return $False
        }
        if (get-module -ListAvailable -Name TrellOps){return $True}    
    }
}


Function LookForFile (){
    $files = Get-ChildItem | where {$_.Extension -eq ".xls"}
    if(!$files){
        msgBox "No .xls file!"
        return $False
    }elseif($files.count -eq 1){
        return $files
    }else{ 
        msgBox "Too many .xls files!"
        return $False
    }
}

Function ExportWSToCSV ($excelFile)
{
    try{
        $E = New-Object -ComObject Excel.Application
    }catch{
        msgBox "Couldn't open Excel. Is it installed?"
        return $False
    }
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($excelFile.FullName).Worksheets
    if(!$wb){
        msgBox "No worksheets!"
        $n = $False
    }elseif($wb.count -eq 1){
        $n = $pwd.path + "\"+$excelFile.BaseName + "_" + $wb[1].Name + ".csv"
        $wb[1].SaveAs($n, 6)
    }else{
        msgBox "Too many worksheets!"
        $n = $False
    }
    $E.Quit()
    return $n
}

Function ReadReservations($csvFile){
    $Null = @(
        $csv = Import-Csv $csvFile
        $reservations = New-Object System.Collections.ArrayList

        foreach ($res in $csv | Group-Object reservation_id){

            $items = New-Object System.Collections.ArrayList   
            foreach ($item in $res.Group){
                $name = $item.cust_name
                $code = $item.itemcode #or equip_code?
                $type = $item.itemtype
                $make = $item.itemmake
                $model = $item.itemname
                $itemObj = New-Object -TypeName PSObject -prop (@{
                    'Name' = $name;
                    'Code' = $code;
                    'Type' = $type;
                    'Make' = $make;
                    'Model' = $model;
                })
                $items.Add($itemObj)
            }#end loop

            $id = $res.Group[0].reservation_id
            $name = $res.Group[0].contact_lname + ", " + $res.Group[0].contact_fname
            $location = $res.Group[0].accommodations
            $roomNumber = $res.Group[0].room_num
            $ready = $res.Group[0].ret_ready
            $notes = $res.Group[0].ret_notes
            $van = $res.Group[0].ret_van
            $reservation = New-Object -TypeName PSObject -Prop (@{
                'ID' = $id;
                'Name' = $name;
                'Location' = $location;
                'RoomNumber' = $roomNumber;
                'Ready' = $ready;
                'Notes' = $notes;
                'Van' = $van;
                'Items' = $items;
            })
            $reservations.Add($reservation)
        }#end loop
    )#end null
    return $reservations
}#end function



Function AuthenticateTrello(){
    $Null = @(
        try{
            $auth = Import-Clixml "dontDelete.data"
        }catch{
            Write-Host "dontDelete.data not found"
            $key = "1dbf1617b88f876120044a53c04c4308"     
            $auth = New-TrelloToken -Key $key -AppName "Trello Inputr" -Expiration "never" -Scope 'read,write'
            $auth | Export-Clixml "dontDelete.data"
        }
        try{
            Get-TrelloBoard -Token $auth -ErrorAction Stop
        }catch{
            MsgBox $Error[-1].Exception
            $auth = $False
        }
    )
    return $auth
}
Function ClearTheBoard($Auth){
    Write-Host "Clearing the board"
    $Null = @(       
        $cards = Get-TrelloCard -Token $auth -Id $board.id -List $listReady
        $cards += Get-TrelloCard -Token $auth -Id $board.id -List $listPending
        $cards += Get-TrelloCard -Token $auth -Id $board.id -List $LeChamoisPending
        $cards += Get-TrelloCard -Token $auth -Id $board.id -List $AavaPending
        $cards += Get-TrelloCard -Token $auth -Id $board.id -List $preAssigned

        if($cards){
            $progress = 0
            Write-Progress -Activity "Deleting old cards..." -PercentComplete $progress -Status "Starting"
            foreach ($card in $cards){
                $progress += 100/$cards.count
                Write-Progress -id 1 -Activity "Deleteing old cards..." -PercentComplete $progress -Status $card.Name
                Remove-TrelloCard -Token $auth -Id $card.id

            }
        }
    )
    return $True
}


Function CreateCustomChecklist($Items){
    $checklist = Add-TrelloChecklist -Token $Auth -Id $Card.Id -Name "Equiptment" -Position bottom
    $progress = 0
    foreach ($item in $items){
        if ($item.Code.ToString() -ne ""){
            $string = $item.Name + " - " + $item.Type + " - " + $item.Code.ToString()
        }else{
            $string = $item.Name + " - " + $item.Type
        }
        if ($item.Make -ne ""){
            $string += " - " + $item.Make + " " + $item.Model
        }
        $progress += 100/$items.Count
        Write-Progress -ParentId 1 -Id 2 -Activity "Creating Checklist" -Status $string -PercentComplete $progress
        Add-TrelloChecklistItem -Token $Auth -Id $checklist.id -Name $string
    }
}
Function CreateGenericChecklist(){
    $checklist = Add-TrelloChecklist -Token $Auth -Id $Card.Id -Name "Return Log" -Position bottom
    Add-TrelloChecklistItem -Token $Auth -Id $checklist.id -Name "Full Return"
    Add-TrelloChecklistItem -Token $Auth -Id $checklist.id -Name "Partial Return"
    Add-TrelloChecklistItem -Token $Auth -Id $checklist.id -Name "Logged in DSR"
}

Function FindList($listName){
    $list = Get-TrelloList -Token $auth -Id $board.id -Name $listName
    if (!$list){
        MsgBox "Could not find list $($listName)"
        return $False
    }else{
        return $list
    }
}

function GetListAndLabel(){
    if($res.Ready -eq "TRUE"){
        $list = $listReady
        $label = $labelReady
    }else{
        $list = $listPending
        $label = $labelPending
    }

    if($res.Location -eq "1. LE CHAMOIS SKI SHOP"){
        $list = $LeChamoisPending
    }elseif($res.Location -eq "Aava Whistler Hotel"){
        $list = $AavaPending
    }elseif($res.Van){
        $list = $preAssigned        
    }

    return $list, $label
}

Function Export2Trello($Reservations){
    $Null = @(
        $auth = AuthenticateTrello
        if($auth){
            $board = Get-TrelloBoard -Token $auth -Name $boardName
            if (!$board){
                MsgBox "Could not find board $($boardName)"
                return $False
            }
            #lists
            $listReady = FindList -listName $listReadyName
            $listPending = FindList -listName $listPendingName
            $LeChamoisPending = FindList -listName $LeChamoisPendingName
            $AavaPending = FindList -listName $AavaPendingName
            $preAssigned = FindList -listName $preAssignedName

            if(ClearTheBoard -Auth $auth){
                #labels
                $labelReady = Get-TrelloLabel -Token $auth -Id $board.id -Name "Ready"
                $labelPending = Get-TrelloLabel -Token $auth -Id $board.id -Name "Pending"
            
            
                $progress = 0
                
                Write-Progress -id 1 -Activity "Creating cards..." -PercentComplete $progress -Status "Starting"

                foreach ($res in $Reservations){
                    if(!$res.RoomNumber){
                        $name = $res.Name + ", " + $res.Location #no room number
                    }else{
                        $name = $res.Name + ", " + $res.Location + ", #" + $res.RoomNumber 
                    }
                    Write-Host "Creating card: $($name)"
                
                    $progress += 100/$Reservations.Count
                    Write-Progress -Id 1 -Activity "Creating Cards ..." -PercentComplete $progress -Status $name


                    $description = "**Guest Accomodation:** " + $res.Location
                    $description += "`n`n"
                    if ($res.RoomNumber -ne ""){
                        $description += "**Room Number:** " + $res.RoomNumber
                        $description += "`n`n"
                    }
                    $description += "**Items to be Collected:** " + $res.Items.Count.ToString()
                    $description += "`n`n"
                    $description += "**Storeage Location:** " #+ Not sure where this info is?
                    $description += "`n`n"
                    $description += "**Notes:** " + $res.Notes
                    if ($res.Van -ne ""){
                        $description += "`n`n"
                        $description += "**Van:** " + $res.Van
                    }

                    $list,$label = getListAndLabel
                    $card = New-TrelloCard -Token $auth -id $list.id -Name $name -Description $description -Label $label.id -Position bottom                    
                    CreateCustomChecklist -Items $res.Items 
                    CreateGenericChecklist
                }#end loop
            }else{ 
                Write-Host "Failed to clear the board"
                return $False
            }
        }else{            
            return $False
        }

    )#end null
    return $True

}
Function CleanUp($csv = $False, $xls = $False){
    #delete csv and .xls
    if ($csv){
        Remove-Item -Path $csv
    }
    if ($xls){
        Remove-Item -Path $xls.FullName
    }
}



$boardName = "Returns"
$listReadyName = "Returns Ready"
$listPendingName = "Returns Pending"
$LeChamoisPendingName = "Returns Le Chamois Pending"
$AavaPendingName = "Returns Aava Pending"
$preAssignedName = "Returns Preassigned"

"Making sure TrellOps is installed"
if (CheckForTrellOps){
    "Looking for file"
    $file = LookForFile
    if ($file) {
        "Converting to csv"
        $csvFile = ExportWSToCSV -excelFile $file
        if ($csvFile){
            "Reading csv"
            $res = ReadReservations -csvFile $csvFile
            "Writing to trello"
            $ran = Export2Trello -Reservations $res
            if ($ran){
                "Cleaning up"
                CleanUp -csv $csvFile #-xls $file
                MsgBox -Text "Finished!" -Err "Information"
            }else{Read-Host -Prompt “Didn't ran”}
        }else{Read-Host -Prompt "Didn't run"}
    }else{Read-Host -Prompt "Didn't run"}
}else{Read-Host -Prompt "couldn't install"}

