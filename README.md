# Teams-Call-Queue-Requests

Contact: [LinkedIn](https://www.linkedin.com/in/mark-webb-797aba69/) 

This document will give you an overview on how to create the solution and add it to Microsoft Teams. Some of the elements will almost certainly need to be tweaked based on your needs and when you import the Power App solution you'll need to authenticate the Power Automate flows using an account(s) that have permissions to the different elements (SharePoint, Exchange etc). I've tried to document below what each of these are so make sure all the prerequisites are done before you try importing the solution Power Apps.

In the same way with the Powershell scripts that are in the library, 99% can be simpy copied and pasted but there are a couple parts that you may need to change based on how you want to implement the solution. I've tried to add comments into each script where this might apply. In addition if you're trying to build on top of this version (such as adding additional settings to a Call Queue that aren't in this version, SLA's for example) then you'll need to update the scripts to add those elements in (as well as the Power App, Power Automate Flows, Sharepoint lists). I'm more than happy to give advice if you're trying to add new elements and aren't sure how something works or where something needs changing. Message me on LinkedIn using details above.

# Setup

## SharePoint

You'll need to setup 5 SharePoint lists...

+ Call Queues
+ Call Queue Requests
+ Call Queue Amendments
+ Call Queue Agent Amendments
+ Call Queue Deletions

I've added csv files with the schema attached to _this_ folder. Create a new SharePoint list using each csv file.

You'll need a service account that Power Automate and the Logic App will use to read and write to these lists. You _could_ set the Power App/Power Automate to use the logged in users credentials to do this but that then needs you to give read/write permissions for the lists to every user which we don't want. So create an account in M365, you can call it what you want, and then give it permissions to each SharePoint list. Again you could give the account Edit permissions to the entire SharePoint site but to scope permissions to only what's required it's better to break the permission inheritence on each list and then grant Full Control of the list(s) to the service account.

## Entra Custom Role

You need to create a custom role in Entra which will have the permissions we need to assign to the Automation Account and Logic App so they can create/amend/delete M365 Groups and Resource Accounts. 

This isn't possible using either the Entra GUI or PowerShell as far as I can tell, almost certainly because the microsoft.directory/users/create and microsoft.directory/users/delete permissions are Privileged permissions. So you need to do this using Graph Explorer. Sign in to Graph Explorer with an account that has permissiont to create a Custom Entra role and use the following parameters

Type: Post

URI: https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions

Body:
```
{
  "description": "Custom Role - Teams Call Queue Creation. Used by Logic App and Automation Account to create and managed Teams Call Queues",
  "displayName": "Custom Role - Teams Call Queue Creation",
  "rolePermissions": [
    {
      "allowedResourceActions": [
        "microsoft.directory/users/create",
        "microsoft.directory/users/delete",
        "microsoft.directory/groups.unified/create",
        "microsoft.directory/groups.unified/delete",
        "microsoft.directory/groups.unified/members/update"
      ]
    }
  ],
  "isEnabled": true
}
```

You should get a 201 response like this

![image](https://github.com/user-attachments/assets/3efcd8db-01ef-4c07-94be-d104aca5a26c)

## Resource Group

If you haven't already got an Azure Resource Group setup for this, set one up or use an existing one.

Create a new Automation Account (or use an existing) and add the PowerShell scripts found _here_. You can use either PowerShell 7 or PowerShell 5.1, dealers choice ;)

## Automation Account

Make sure you've enabled a System Assigned managed identity for the Automation Account via the Identity blade in the Automation Account.

The Deletion and Amendment scripts shouldn't need to be changed at all from how they've been uploaded but for the creation scripts, you'll need to edit this line in each script to specify the domain name for the Resource Account

$CQUPN = $CallQueueUPN + "@xxx.onmicrosoft.com"

Alternatively you could add your domain name as an Automation variable, import that at the beginning of the script and use that value.

Add the following modules to the Automation Account

+ Microsoft Teams (used for Creating/Amending/Deleting Call Queues)
+ Microsoft.Graph.Authentication (used for connecting to Graph)
+ Microsoft.Graph.Users (used for deleting/amending Resource Accounts)
+ PnP.PowerShell (used for writing data to SharePoint)

Assign the Custom Role you created earlier the Automation Account in Entra. Search using the Automation Account name and it should show up as an Enterprise Application.

![image](https://github.com/user-attachments/assets/c8cfbd4e-86b6-4be0-9ccd-5cce219898ba)

make sure it's added as an Active assignment, not just Eligible.

![image](https://github.com/user-attachments/assets/12e3fbfc-4631-49e2-a069-b73bb2633171)


