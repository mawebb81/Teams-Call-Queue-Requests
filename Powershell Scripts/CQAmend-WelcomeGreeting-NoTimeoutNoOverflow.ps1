Param (

    [Parameter(Mandatory=$true)]

    [String]

    $CallQueueID,


    [Parameter(Mandatory=$true)]

    [String]

    $CallQueueName,


    [Parameter(Mandatory=$true)]

    [String]

    $TimeoutThreshold,
 

    [Parameter(Mandatory=$true)]

    [String]

    $OverflowThreshold,

 

    [Parameter(Mandatory=$true)]

    [String]

    $RoutingMethod,

 

    [Parameter(Mandatory=$true)]

    [String]

    $PresenceBasedRouting,



    [Parameter(Mandatory=$true)]

    [String]

    $AllowOptOut,

 

    [Parameter(Mandatory=$false)]

    [string]

    $ConferenceMode,

 

    [Parameter(Mandatory=$true)]

    [String]

    $AgentAlertTime,

 

    [Parameter(Mandatory=$false)]

    [String]

    $WelcomeTextToSpeech,

    
    [Parameter(Mandatory=$true)]

    [String]

    $GroupID,

    [Parameter(Mandatory=$true)]

    [String]

    $ResourceID,

    [Parameter(Mandatory=$true)]

    [String]

    $NameChange

)


Connect-MicrosoftTeams -Identity

Connect-MgGraph -Identity

##Update display name if needed##

If($NameChange -eq "Yes") {

Try {

    Update-MgUser -UserID $ResourceID -DisplayName $CallQueueName -ErrorAction Stop
}

Catch {

    Throw "Error changing resource account display name, the error is $($Error[0])"
}

Write-Output "Changed CQ Display Name succesfully"

}

##Replace any spaces or invalid characters in CQ Name as it's used for UPN##
 
$CharactersToRemove = "[,' ./\\\(\)\{\}\[\]]"

#$CallQueueNameNS = $CallQueueName -replace '\s',''

##Replace any spaces in Routing Method as needs to be all one word##

$RoutingMethodNS = $RoutingMethod -replace $CharactersToRemove,''

##Convert values for Presence Based Routing and Opt Out to Boolean from our text values## 

$ConvertedPresRouting = Switch ($PresenceBasedRouting) {'yes'{$True}'No'{$False}}

 
$ConvertedAllowOptOut = Switch ($AllowOptOut) {'yes'{$True}'No'{$False}}

Write-Output "Amending Call Queue"

##Amend call queue##

Try {

#$CQID = (Get-CSCallQueue -NameFilter $CallQueueName).Identity

Set-CsCallQueue -Identity $CallQueueID -OverflowThreshold $OverflowThreshold -TimeoutThreshold $TimeoutThreshold -RoutingMethod $RoutingMethodNS -PresenceBasedRouting $ConvertedPresRouting -UseDefaultMusicOnHold $True -DistributionLists $GroupID -AgentAlertTime $AgentAlertTime -AllowOptOut $ConvertedAllowOptOut -WelcomeTextToSpeechPrompt $WelcomeTextToSpeech -OverflowAction DisconnectWithBusy -OverflowDisconnectTextToSpeechPrompt $null -TimeoutAction Disconnect -TimeoutDisconnectTextToSpeechPrompt $null -ConferenceMode $True -LanguageID "en-GB" -ErrorAction Stop

}

Catch {

    Throw "Errors amending Call Queue, error is $($error[0]) exiting script"

}



Write-Output "Call Queue amended, finishing up"

