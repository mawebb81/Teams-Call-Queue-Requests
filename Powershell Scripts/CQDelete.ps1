Param (

    [Parameter(Mandatory=$true)]

    [String]

    $CallQueueID,

    [Parameter(Mandatory=$true)]

    [String]

    $ResourceAccountID

)

##Connect to Teams and Graph using Managed Identity##

Connect-MicrosoftTeams -Identity

Connect-MgGraph -Identity

##Remove Resource Account from Call Queue so CQ can be deleted. CQ would fail to delete if Resource Account still associated##

Try {

    Remove-CsOnlineApplicationInstanceAssociation -Identities $ResourceAccountID -ErrorAction Stop
}

Catch {

    Throw "Error removing Resource Account association error is $($error[0])" 
}

##Wait for replication##

Start-Sleep -Seconds 30

##Delete Call Queue##

Try {

Remove-CsCallQueue -Identity $CallQueueID -ErrorAction Stop

}

Catch {

    Throw "Error deleting call queue, error is $($error[0]) exiting"
}

Write-Output "Call Queue Deleted succesfully"

##Delete Resource Account##

Try {

Remove-MgUser -UserId $ResourceAccountID

}

Catch {

    Throw "Error deleting Resource Account, error is $($error[0]) exiting"
}

Write-Output "Resource Accoint deleted succesfully"
