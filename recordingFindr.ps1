#Alex Dobrovansky
#05 Dec 17

#v3

$matchList = New-Object System.Collections.ArrayList
$pv5List = @()


Add-Type -AssemblyName System.Windows.Forms
$pv5Location = New-Object System.Windows.Forms.FolderBrowserDialog -property @{
    ShowNewFolderButton = $False
    Description = "Select location of recordings to search"
}
[void]$pv5Location.ShowDialog()
set-location $pv5Location.SelectedPath



#all the logic:
$PV5s = Get-ChildItem | Where-Object {$_.Extension -eq ".pv5"}
foreach ($file in $PV5s){
    [xml]$xml = [regex]::Match($(get-content $file),"<XML>.+<\/XML>").Value
    #time

    $callDuration = [datetime]$xml.XML.End - [datetime]$xml.XML.Start

    $info = New-Object -TypeName PSObject -Property (@{
        "Name" = $file.name;
        "ID" = $xml.XML.Id;
        "StartTime" = $xml.XML.Start;
        "CLI" = $xml.XML.CLI;
        "DDI" = $xml.XML.DDI;
        "Direction" = $xml.XML.DIRECTION;
        "Duration" = $callDuration;
    })
    
    $pv5List += $info

}
$pv5List | Out-GridView -Title "RecordingFindr" -PassThru | foreach { $matchList.Add($_) }

$matchList

$options = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes", "&No")
$copyMatchedRecordings = $Host.UI.PromptForChoice("Copy Matches","Would you like to copy the selected matches?", $options, 1 ) 


if($copyMatchedRecordings -eq 0 -and $matchList){
    Add-Type -AssemblyName System.Windows.Forms
    $copyLocation = New-Object System.Windows.Forms.FolderBrowserDialog -property @{ Description = "Select location to copy recordings to" }
    [void]$copyLocation.ShowDialog()
    foreach($pv5 in $matchList){ Copy-Item -Path $pv5 -Destination $copyLocation }
}
