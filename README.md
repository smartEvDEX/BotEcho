# Transcription bot

This bot is created for processing the transcriptions of a Teams Meeting and then give you a summary of it. It only accept files of txt and vtt type.

## How to get the transcripts?

To get the transcripts of a meeting you must activate them when it starts. To activate them you must go to more actions (the three points) and search for "Start transcripts". The transcripts are generated and at the end of the meeting you can download them. They are only available in English.

## How to use the bot? 

- Download the project from github. 
- Register the bot in the Azure portal. To register it you must follow the steps indicated in the documentation.   
- Modify the manifest.json with the bot id registered in the Azure portal. You have to change the application id, the webApplicationInfo id and the bot id. 
- Install the bot in Microsoft Teams. To do this you must go to Microsoft Teams > Applications > Load a custom application > add the .zip of the manifest. The manifest is a zip containing a file named color.png, the outline.png and the manifest.json. You can modify the images if you wish, but they must follow the following conditions. 
- In case you don't accept the manifest you can check it at the following link: https://dev.teams.microsoft.com/appvalidation.html  


## How to register the bot? 

- Go to the Azure portal.  
- In the right pane, select Create a resource. 
- In the search box, type bot and press Enter. 
- Select the Azure Bot card. 
- Select Create. 
- Type the required values. Remember that you can change the Rate Plan to Free. 
- Select Review + Create. 
- If the validation passes, select Create. 
- Select Go to resource group. You should see the bot and related Azure Key Vault resources in the resource group you selected. 
In case you have further questions you can refer to the Azure documentation at the attached link. 
https://docs.microsoft.com/es-es/azure/bot-service/bot-service-quickstart-registration?view=azure-bot-service-4.0&tabs=csharp%2Ccshap

## How to use the bot locally? 

To test the bot locally you have to install the Bot Framework Emulator application.  
Start the Bot Framework Emulator application. 
Click on Create a new bot configuration. 
Put the name you want to the bot, endpoint URL: https://localahost:3978/api/messages, and fill in the corresponding Microsoft App Id and password. Click Save and connect. 
Open the bot project in Visual Studio Code, and run the project by clicking Run > Star Debugging. Run the project with node.js, in case you don't have it, install it from this link: https://nodejs.org/es/download/  
