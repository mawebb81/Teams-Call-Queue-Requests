Param (

    [Parameter(Mandatory=$true)]

    [String]

    $CallQueueDisplayName,


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
    $M365GroupID,


    [Parameter(Mandatory=$true)]

    [String]

    $OverflowTextToSpeech,

    [Parameter(Mandatory=$true)]

    [String]
    $CallQueueUPN

)


Connect-MicrosoftTeams -Identity

##Get Group ID##

$GroupID = $M365GroupID

##Replace any spaces in Routing Method as needs to be all one word##

$RoutingMethodNS = $RoutingMethod -replace '\s',''

##Set UPN for CQ##

## REPLACE xxx.onmicrosoft.com with your tenant domain, e.g. contoso.com ##

$CQUPN = $CallQueueUPN + "@xxx.onmicrosoft.com"

##Convert values for Presence Based Routing and Opt Out to Boolean from our text values## 

$ConvertedPresRouting = Switch ($PresenceBasedRouting) {'yes'{$True}'No'{$False}}

 
$ConvertedAllowOptOut = Switch ($AllowOptOut) {'yes'{$True}'No'{$False}}

 

Write-Output "Creating new resource account for Call Queue"

Try{

$AppInstance = New-CsOnlineApplicationInstance -UserPrincipalName $CQUPN -ApplicationId "11cd3e2e-fccb-42ad-ad00-878b93575e07" -DisplayName $CallQueueDisplayName -ErrorAction Stop

}

Catch {

    Throw "Error creating resource account, error is $($error[0]) exiting script"

}
 

Write-Output "Created Resource account, waiting 30 seconds for replication"

 
Write-Output "Creating Call Queue"

Try{

$NewCallQueue = New-CsCallQueue -Name $CallQueueDisplayName -OverflowThreshold $OverflowThreshold -TimeoutThreshold $TimeoutThreshold -RoutingMethod $RoutingMethodNS -PresenceBasedRouting $ConvertedPresRouting -UseDefaultMusicOnHold $True -DistributionLists $GroupID -AgentAlertTime $AgentAlertTime -AllowOptOut $ConvertedAllowOptOut -WelcomeTextToSpeechPrompt $WelcomeTextToSpeech -OverflowAction "DisconnectWithBusy" -OverflowDisconnectTextToSpeechPrompt $OverflowTextToSpeech -ConferenceMode $True -LanguageID "en-GB" -ErrorAction Stop

}

Catch {

    Throw "Error creating Call Queue, error is $($error[0]) exiting script"

}
 

Start-Sleep -Seconds 240

$AvailableNumbers = Get-CsPhoneNumberAssignment -ActivationState Activated -CapabilitiesContain VoiceApplicationAssignment -PstnAssignmentStatus Unassigned | Select TelephoneNumber

$RandoNumber = $AvailableNumbers | Get-Random

Try{

Set-CsPhoneNumberAssignment -Identity $CQUPN -PhoneNumber $RandoNumber.TelephoneNumber -PhoneNumberType CallingPlan -ErrorAction Stop

}

Catch {

    Throw "Error assigning phone number to Call Queue"

}
 

Write-Output "Call Queue created, moving on to Associating resource account to Call Queue"

$AssociationComplete = $False

[int]$Retry = "1"

Do {

Try {

    New-CsOnlineApplicationInstanceAssociation -ConfigurationId $NewCallQueue.Identity -Identities $AppInstance.objectId -ConfigurationType CallQueue -ErrorAction Stop

    $AssociationComplete = $True

}

Catch {

    If ($Retry -ge 10) {

    "Error associating Resource Account with Call Queue, error is $($error[0])"

    $AssociationComplete = $True

    }

    Else {

    Write-Output "Failed to associate account, trying again"

    Start-Sleep -Seconds 10

    $Retry++

    }

}

}

While ($AssociationComplete -eq $False)

Write-Output "Queue and Resource Account associated succesfully"
$CQDetails = [PSCustomObject]@{
    ResourceAccID = $AppInstance.objectId
    CQID = $NewCallQueue.Identity
    TelephoneNumber = $RandoNumber.TelephoneNumber
}
Write-Output ($CQDetails | ConvertTo-Json)