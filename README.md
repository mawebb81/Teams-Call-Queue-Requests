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

## 
