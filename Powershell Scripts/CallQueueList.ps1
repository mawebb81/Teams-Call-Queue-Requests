##Set variables to connect to SharePoint and Graph##

 

$PNPCert = Get-AutomationVariable -Name PnP-Thumbprint

 

$PNPAppID = Get-AutomationVariable -Name PNMS-AppID

 

##Connect to Graph for Teams query and SharePoint to write to list/documents##





Connect-MgGraph -Identity

Get-MgUser -UserId 71a83893-86ff-4cb6-b358-5f760a132bea

Connect-MicrosoftTeams -Identity

$Connection = Connect-PnPOnline -ClientId $PNPAppID -Url https://ukhodev.sharepoint.com/sites/VV-CallQueues -Thumbprint $PNPCert -Tenant "ukhodev.onmicrosoft.com" -ReturnConnection


##Connect to Teams and Graph##



 

##Get all existing list items##

 

$ListItems = Get-PnPListItem -List "Call Queues" -Connection $Connection

 

#remove all current list items so we can refresh/add up to date items##

 

ForEach ($item in $ListItems) {

 

    Remove-PnPListItem -List "Call Queues" -Identity $item.Id -Force -Connection $Connection

 

}

 
# Initialize variables for Call Queue variable##
$AllCallQueues = @()
$batchSize = 100
$skip = 0

# Loop to get all call queue

Try{
    
do {
    # Get a batch of call queues
    $callQueues = Get-CsCallQueue -First $batchSize -Skip $skip -ErrorAction Stop

    # Add the batch to the allCallQueues array
    $allCallQueues += $callQueues

    # Update the skip value for the next batch
    $skip += $batchSize
} while ($callQueues.Count -eq $batchSize)

}

Catch {
    Throw "Failed to get Call Queues...exiting..."
}



##Write details to SharePoint List. First, need to convert the agents property to a string as it's currently an object and won't write to list in that state when multiple users in agents property##

 

ForEach ($Queue in $CallQueues) {

 

    Try {

 

        $UPN = $Queue.Name + "@ukhodev.onmicrosoft.com"

 

        ##Convert AAD ObjectID's contained in Agents property to UPN's##

 

        $Agents = $Queue.Agents

 

        ForEach ($agent in $agents) {

 

            $AgentID = (Get-MgUser -UserId $agent.ObjectId)

 

            $agent.objectID = $agentID.UserPrincipalName + ","

        }

 

         $AgentString = $queue.agents.objectID | Out-String -NoNewline

 

        

        ##Convert Routing Method to named value as default writes to SharePoint as an integer for some reason. If ever work out why can change this.##

 

        If ($queue.RoutingMethod -eq 0) {

            $QRoutingMethod = "Attendant"

        }

 

        elseIf ($queue.RoutingMethod -eq 1) {

            $QRoutingMethod = "Serial"

        }

 

        elseIf ($queue.RoutingMethod -eq 2) {

            $QRoutingMethod = "Round Robin"

        }

 

        elseIf ($Queue.RoutingMethod -eq 3) {

            $QoutingMethod = "Longest Idle"

        }

 

        ##Get details of welcome greeting if there is one. If length of text 1 or greater, store message. If 0 then no greeting##

 

        If ($Queue.WelcomeTextToSpeechPrompt.Length -ge 1) {

 

            $WelcomeGreeting = "Yes"

 

        } Else {

 

            $WelcomeGreeting = "No"

        }

 

        ##Do same for Timeout greeting##

 

        If ($Queue.TimeoutDisconnectTextToSpeechPrompt.Length -ge 1) {

 

            $TimeoutGreeting = "Yes"

 

        } Else {

 

            $TimeoutGreeting = "No"

        }

 

        ##Do same again for Overflow greeting##

 

        If ($Queue.OverflowDisconnectTextToSpeechPrompt.Length -ge 1) {

 

            $OverflowGreeting = "Yes"

 

        } Else {

 

            $OverflowGreeting = "No"

        }

 

        ##Get Queue Contact/Owner, will return either first direct user assigned or first group member if M365 Group linked to Queue. Need to trim last character as it's a "," based on how we need to present list for powerapp##

 

        If ($Queue.Agents[0].ObjectId.Length -ge 1) {

 

            $TrimID = $Queue.Agents[0].ObjectId.Substring(0,$Queue.Agents[0].ObjectId.Length-1)

 

            $Owner = Get-MgUser -UserId $TrimID | Select UserPrincipalName

 

            $Contact = $Owner.UserPrincipalName

 

        } Else {

 

            $Contact = "No Owner/Agents"

 

        }

 

        ##Get M365 Group ID for Queue, not all have them so set value based on if they do or not##

 

        If ($Queue.DistributionLists.Guid.Length -ge 1) {

 

            $QueueGroupID = $Queue.DistributionLists.Guid

 

        } Else {

 

            $QueueGroupID = "No M365 Group"

        }

 

        $QueueID = $Queue.Identity

 

        $CQResourceID = $queue.ApplicationInstances -join ','

 

        $CQResourceName = Get-CsOnlineUser -Identity $CQResourceID | Select UserPrincipalName, LineUri

 

        If ($CQResourceName.LineUri.Length -ge 1) {

 

            $PhoneNumber = $CQResourceName.LineUri.Substring(4)

 

        }

 

        Else {

 

            $PhoneNumber = "No Number"

 

        }

 

        ##Convert values of Opt Out and PBR to Yes/No rather than True/False##

 

        $ConvertedPresRouting = Switch ($Queue.PresenceBasedRouting) {$True{'Yes'}$False{'No'}}

 

        $ConvertedOptOut = Switch ($Queue.AllowOptOut) {$True{'Yes'}$False{'No'}}

 

        ##Add each item to SharePoint##

 

        Add-PnPListItem -List "Call Queues" -Value @{"Name" = $Queue.Name; "WelcomeGreeting" = $WelcomeGreeting; "RoutingMethod" = $QRoutingMethod; "Agents" = $AgentString; "AllowOptOut" = $ConvertedOptOut; "ConferenceMode" = $Queue.ConferenceMode; "AgentAlertTime" = $Queue.AgentAlertTime; "OverflowThreshold" = $Queue.OverflowThreshold; "TimeoutThreshold" = $Queue.TimeoutThreshold; "WelcomeTexttoSpeech" = $Queue.WelcomeTextToSpeechPrompt; "PresenceBasedRouting" = $ConvertedPresRouting; "CallQueueNumber" = $PhoneNumber;"OverflowGreeting" = $OverflowGreeting; "TimeoutGreeting" = $TimeoutGreeting; "TimeoutText" = $Queue.TimeoutDisconnectTextToSpeechPrompt; "OverflowText" = $Queue.OverflowDisconnectTextToSpeechPrompt; "QueueContact" = $Contact; "M365GroupID" = $QueueGroupID; "QueueID" = $QueueID; "ResourceAccName" = $CQResourceName.UserPrincipalName; "ResourceAccID" = $CQResourceID} -Connection $Connection -ErrorAction Stop

 

    } Catch {

 

        Write-Output "error is $($error[0])"

 

    }

}