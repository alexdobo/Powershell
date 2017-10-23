#Simple function to send a slack message
#Requires a webhook URL provided by slack. It will only send messages to the channel the webhook is linked to

function Send-SlackMsg {
    param($Username,$Body)

    $message = @{
        username = $Username
        text = $Body
    }
    $json = $message | ConvertTo-Json

    $webhookURL = "https://hooks.slack.com/services/" #insert your webhook URL here
    Invoke-RestMethod -Method Post -Uri $webhookURL -Body $json
}
