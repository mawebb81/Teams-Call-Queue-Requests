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

Create a new Automation Account (or use an existing) and add the PowerShell scripts found [here](docs/Powershell Scripts). You can use either PowerShell 7 or PowerShell 5.1, dealers choice ;)

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


## Logic Apps

Now you need to create the 4 logic apps that will handle the requests that are added to SharePoint.

You need one for each request type

+ Call Queue Requests
+ Call Queue Deletions
+ Call Queue Amendments
+ Call Queue Agent Amendments

Create them using whatever your required parameters are for resource location, whether they are consumption or not etc

Once created, before configuring any of the steps, enable the System Assigned managed identity for each via the Identity blade in the Logic App

![image](https://github.com/user-attachments/assets/98587a8a-34a3-458a-b711-dcb71dbed958)

Once you've enabled that, go back into your Automation Account and grant each of the Logic App managed identities the Automation Contributor role so the logic app(s) can call and read automation jobs. 
![image](https://github.com/user-attachments/assets/67d0fe1e-55b6-464d-b461-334369532dfa)

![image](https://github.com/user-attachments/assets/3eb32ee0-6571-40cf-90e7-1ca2b375cd3e)

Now, like you did for the Automation Account, grant the System Assigned managed identity of each Logic App the Custom Entra role you created earlier. Same process as above.

You can now start builing your Logic Apps. When creating these from scatch I'd recommend creating the Deletion Logic App first as this contains all the different connections we need across all of the Logic Apps and means we can create the remaining ones using code rather than having to build them nanually.

Below is the structure of the Deletion Logic App

![image](https://github.com/user-attachments/assets/3b87255d-f03d-456b-8ec7-f93d2d1ccf00)

You'll also need to create two parameters for this Logic App:

+ SharePoint Site
+ List Name

Both of these are string variables. The first is the URL of your SharePoint site where your lists are and the second is the name of the list where your deletion requests are.

I've listed below each of the steps for this Logic App, add each one with the relevant info you need in your environment

1. When a new item is created (trigger)
   - Site Address = SharePoint Site parameter
   - List Name = List parameter
   - How often do you want to check for items = whatever value you wish
   - Connection = Your service account with permission to the SharePoint list
2. Get changes for an item or file (properties only)
   - Site Address = SharePoint Site parameter
   - List Name = List parameter
   - Id = ID of the SharePoint list item (dynamic value)
   - Since = Trigger Window Start Token (dynamic value)
   - Connection = Your service account with permission to the SharePoint list
3. Initialize Variable
   - Type = String
   - Value = substring(triggerBody()?['PhoneNumber'],1,12) (this will depend on how your number is formatted in your list, this example strips the first character as it has a '+' which we don't want.
4. M365 Group HTTP Action (use the 
   - URI = https://graph.microsoft.com/v1.0/groups/_M365GroupID_ (the M365 Group ID is the value from your SharePoint list)
   - Method = Delete
   - Connection = System Assigned Managed Identity
5. Create Automation Job
   - Subscription = _your Azure subscription_
   - Resource Group = _your Azure Resource Group_
   - Automation Account = _your Automation Account_
   - Runbook Name = CQDelete (or whatever the name of the runbook is you have for deleting Call Queues
   - Wait for Job = Yes
   - CallQueueID parameter = _Call Queue ID_ (this is the value from the SharePoint list
   - ResourceAccountID parameter = _Resource Account ID_ (this is the value from the SharePoint list)
   - Connection = System Assigned Managed Identity
6. Condition
   - Expression = Runbook Status is equal to Completed
7. True Path
   - Site Address = SharePoint Site parameter
     - List Name = List parameter
     - Id = ID of the SharePoint list item (dynamic value)
     - Status column = Completed
     - Connection = Your service account with permission to the SharePoint list
   - Delete SharePoint Item
     - Site Address = SharePoint Site parameter
     - List Name = Call Queues (or whatever name you have given your list that contains all your Call Queues)
     - Id = _SharePoint List ID_ (this is the value from the SharePoint list)
     - Connection = Your service account with permission to the SharePoint list
   - Send Email
     - To - _Requester Email_ (this is the value from the SharePoint list)
     - Subject - _Whatever you want_
     - Body - _Whatever you want_
     - Connection = _Whatever account you want the email to come from_
8. False Path
   - Site Address = SharePoint Site parameter
     - List Name = List parameter
     - Id = ID of the SharePoint list item (dynamic value)
     - Status column = Failed
     - Connection = Your service account with permission to the SharePoint list
   - Send Email
     - To - _Requester Email_ (this is the value from the SharePoint list)
     - Subject - _Whatever you want_
     - Body - _Whatever you want_
     - Connection = _Whatever account you want the email to come from_

Now we have this one created, we can create the rest using the code below and just make a few changes to the code, this is mainly to update the connection references and parameters to reflect your environment.

The code for the other 3 Logic Apps can be found _here_.

For each one open the code in whatever your preferred code editor is i.e. VS Code etc and then at the bottom where you see the connections specified update the sections for each that are highlighted below
![image](https://github.com/user-attachments/assets/28e25e38-a444-495c-a26e-c1f38860ae57)

You won't have all these for each Logic App but the principle is the same. For each connection change aaa-aaa-aaa to your Azure subscription number, change bbbb to the name of your Azure Resource Group and then for the names of the connections, such as sharepointonline-2, make sure these look exactly the same as your code from the Call Queue Deletion Logic App. I've found it easier to have the code view for the Deletion Logic App on one screen and the code for the new one I'm creating on another so I can easily compare/copy and paste as needed between them.

Next, locate the parameters section of the code and amend the items there to match your environment. The example below shows two parameters, you'll have more depending on the type of Logic App. Update the values as needed, in this example repalcing aaa with the SharePoint Site URL and bbbb with the List Name

![image](https://github.com/user-attachments/assets/03177bb8-f0fc-47ae-8d69-d8fc005fa38d)

Once you've updated all of those you can copy all of the code for the Logic App, return to the Code View in the new Logic App and paste it in. Click Save and you should now have the new Logic App created with the right connections/parameters etc. Go to the Logic App Designer to check each step is valid and you don't have any errors.

## Power App/Power Automate

The last part is to create/import the Power App and Power Automate flows. I'm assuming you've got your own Power App environment to use already of that you've had one created for you to use for this. 

In the environment go to Solutions and click Import Solution. Upload the zip file that can be found _here_. This will import the Power App and the associated Power Automate Flows. You'll be asked to re-establish the connections for the Power Automate flows. Connect using the relevant accounts, for SharePoint connections use the serivce account we gave permissions to the SharePoint lists and for Exchange connections use the same account we used in the Logic App to send emails. 

Once imported the Power App should be ready to go. If you've named the SharePoint lists differently to how I've named them in this example you will need to update the connection references in the Power App and Power Automate to reflect whatever you have named them.

## Add the App to Teams

If you want to add the Power App to Microsoft Teams so it can be used there instead of via a web brower, firstly go to the app in the Power Apps environment and click on the ellipsis icon then Share -> Add to Teams

![image](https://github.com/user-attachments/assets/fac9fad0-e7d7-4f51-a83c-3e681b2065bf)

Add any details you want and then if you're already logged in with an account that has permissions to add apps to Teams click _Add to Teams_ otherwise click _Download app_ and then give the zip file to a Teams Admin to add via Teams Admin Centre.

