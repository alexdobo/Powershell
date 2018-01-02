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
        $board = Get-TrelloBoard -Token $auth -Name $boardName
        if (!$board){
            MsgBox "Could not find board $($boardName)"
            return $False
        }
        $listReady = Get-TrelloList -Token $auth -Id $board.id -Name $listReadyName
        if (!$listReady){
            MsgBox "Could not find list $($listReadyName)"
            return $False
        }
        $listPending = Get-TrelloList -Token $auth -Id $board.id -Name $listPendingName
        if (!$listPending){
            MsgBox "Could not find list $($listPendingName)"
            return $False
        }
        $cardsReady = Get-TrelloCard -Token $auth -Id $board.id -List $listReady
        $cardsPending = Get-TrelloCard -Token $auth -Id $board.id -List $listPending
        $cardsTotal = $cardsReady.Count + $cardsPending.Count
        $progress = 0
        Write-Progress -Activity "Deleting old cards..." -PercentComplete $progress -Status "Starting"
        foreach ($card in $cardsReady){
            $progress += 100/$cardsTotal
            Write-Progress -Activity "Deleteing old cards..." -PercentComplete $progress -Status $card.Name
            Remove-TrelloCard -Token $auth -Id $card.id

        }
        foreach ($card in $cardsPending){
        $progress += 100/$cardsTotal
            Write-Progress -Activity "Deleteing old cards..." -PercentComplete $progress -Status $card.Name
            Remove-TrelloCard -Token $auth -Id $card.id
        }
        
    )
    return $True
}


Function CreateCustomChecklist($Items, $Card, $Auth){
    $checklist = Add-TrelloChecklist -Token $Auth -Id $Card.Id -Name "Equiptment" -Position bottom
    foreach ($item in $items){
        if ($item.Code.ToString() -ne ""){
            $string = $item.Name + " - " + $item.Type + " - " + $item.Code.ToString()
        }else{
            $string = $item.Name + " - " + $item.Type
        }
        
        Add-TrelloChecklistItem -Token $Auth -Id $checklist.id -Name $string
    }
}
Function CreateGenericChecklist($Card, $Auth){
    $checklist = Add-TrelloChecklist -Token $Auth -Id $Card.Id -Name "Return Log" -Position bottom
    Add-TrelloChecklistItem -Token $Auth -Id $checklist.id -Name "Full Return"
    Add-TrelloChecklistItem -Token $Auth -Id $checklist.id -Name "Partial Return"
    Add-TrelloChecklistItem -Token $Auth -Id $checklist.id -Name "Logged in DSR"
}


Function Export2Trello($Reservations){
    $Null = @(
        $auth = AuthenticateTrello
        if($auth){
            if(ClearTheBoard -Auth $auth){
                $board = Get-TrelloBoard -Token $auth -Name $boardName
                if (!$board){
                    MsgBox "Could not find board $($boardName)"
                    return $False
                }
                $listReady = Get-TrelloList -Token $auth -Id $board.id -Name $listReadyName
                if (!$listReady){
                    MsgBox "Could not find list $($listReadyName)"
                    return $False
                }
                $labelReady = Get-TrelloLabel -Token $auth -Id $board.id -Name "Ready"
                $listPending = Get-TrelloList -Token $auth -Id $board.id -Name $listPendingName
                if (!$listPending){
                    MsgBox "Could not find list $($listPendingName)"
                    return $False
                }
                $labelPending = Get-TrelloLabel -Token $auth -Id $board.id -Name "Pending"
            
            
                $progress = 0
                Write-Progress -Activity "Creating cards..." -PercentComplete $progress -Status "Starting"

                foreach ($res in $Reservations){
                    $name = $res.Name + ", " + $res.Location
                    Write-Host "Creating card: $($name)"
                
                    $progress += 100/$Reservations.Count
                    Write-Progress -Activity "Creating Cards ..." -PercentComplete $progress -Status $name


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
                    #write card function

                    if($res.Ready -eq "TRUE"){
                        $card = New-TrelloCard -Token $auth -id $listReady.id -Name $name -Description $description -Label $labelReady.id -Position bottom
                    }else{
                        $card = New-TrelloCard -Token $auth -id $listPending.id -Name $name -Description $description -Label $labelPending.id -Position bottom
                    }
                    CreateCustomChecklist -Items $res.Items -Card $card -Auth $auth
                    CreateGenericChecklist -Card $card -Auth $auth
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
                "Finished!"
                MsgBox -Text "Finished!" -Err "Information"
            }else{Read-Host -Prompt “Didn't ran”}
        }else{Read-Host -Prompt "Didn't run"}
    }else{Read-Host -Prompt "Didn't run"}
}else{Read-Host -Prompt "couldn't install"}
