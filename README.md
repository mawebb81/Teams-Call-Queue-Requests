# Teams-Call-Queue-Requests

Contact: [LinkedIn](https://www.linkedin.com/in/mark-webb-797aba69/) 

This repo is based on a Power App I've created to automate the request and fulfilment of Call Queues within Microsoft Teams. It uses Power Apps to handle the user facing front end to capture info and then SharePoint lists, Logic Apps and Automation Accounts to process the requests and fulfil then. I recently spoke at Commsverse on this and I got a few questions about if the code/project would be available publicly. If you're interested in the link to the Commsverse talk, it's [here](https://events.justattend.com/events/conference-session/247cx437/8d37b3e6).

If you're expecting to be able to copy and paste this into your environment and get going then I'll have to disappoint you slightly, you WILL have to make a few changes to what I've uploaded here. Some of this is just because some references will be to your Azure environment etc so you'll have to update references. Other parts will more likely be you want to change what I've done or your requirements will be slightly different and you want to change stuff. That's the point of making this publicly available, feel free to tweak and change. One of the comments I got when presenting was "Why are you sending emails to people to tell them the Call Queue has been created, why not send a Teams message?".....you can absolutely do that, emails was our choice when we deployed it, if you want to send a Teams message instead, change the Logic App!

The Install Instructions file in this repo should hopefully give you enough detail to guide you through how to set the solution up. It does assume you have a working knowledge of a lot of the building blocks so it's not a full step by step guide.

Hopefully this is useful, if you have any questions or find any issues feel free to message me on LinkedIn on the link above
