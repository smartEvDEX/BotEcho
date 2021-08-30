// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { 
    ActivityHandler,
    TurnContext, 
    MessageFactory,
    TeamsInfo
} = require('botbuilder');

const OpenAI = require('openai-api');
const OpenAITokens = require('openai-nodejs');
const config = require('./config.json');
const officeParser = require('officeparser');
const https = require('https');

const { writeFile } = require('../files/fileSave');
const path = require('path');
const fs = require('fs');

const OPENAI_API_KEY = config.openai.apiKey;
const openai = new OpenAI(OPENAI_API_KEY);
const openai_tokens = new OpenAITokens(OPENAI_API_KEY);

// Variables to prepare the transcripts.
const petition = "Convert my transcription into a first-hand account summary of the meeting: \n"
const resum = "\n Summary:"
const regex = /\d\d:\d\d:\d\d.\d\d\d/;
const regex2 = /^([a-zA-Z]{2,}\s)([a-zA-Z]{2,}\s)([a-zA-Z]{2,})/;
var text = null;
var text2 = null;
var transcription = null;
var finalTranscription = null;
var myArr;

var maxtokens = 150;

// Onine Meeting
// https://docs.microsoft.com/en-us/graph/api/resources/onlinemeeting?view=graph-rest-1.0
// Teams Info
// https://docs.microsoft.com/en-us/javascript/api/botbuilder/teamsinfo?view=botbuilder-ts-latest#getMeetingInfo_TurnContext__string_
// EXAMPLES
// https://github.com/microsoft/botbuilder-samples

// Permisos de la API del BOT
// https://aad.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/CallAnAPI/appId/bacf0ad6-1526-4343-87a1-416f913f946e/isMSAApp/
// Com cridar la API GRAPH REST
// https://docs.microsoft.com/es-es/graph/api/team-list-members?view=graph-rest-1.0&tabs=javascript

class MeetingBot extends ActivityHandler {
    constructor() {
        super();
        // https://github.com/microsoft/BotBuilder-Samples/blob/main/samples/javascript_nodejs/56.teams-file-upload/bots/teamsFileUploadBot.js
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            const attachments = context.activity.attachments;
            
            // Detecta que s'ha adjuntat un arxiu
            if (attachments == undefined) {
                await context.sendActivity("Hola!"); 
                await context.sendActivity("Si m'envies un fitxer amb les transcripcions d'una reunió t'envio el resum."); 

            } else {
                for (let i = 0; i < attachments.length; i++){
                    const file = attachments[i];
                    // En el cas que l'etiquetis en una conversa de grup
                    if(file.contentType == "text/html"){
                        await context.sendActivity("Benvingut!"); 
                        await context.sendActivity("Si m'envies un fitxer per privat amb les transcripcions d'una reunió t'envio el resum.");   
                    } else {
                        transcription = null;
                        finalTranscription = null;
                        text = null;
                        text2 = null;
            
                        // Amb hem de llegim la URL de descarrega i el tipus d'arxiu
                        const downloadUrl = file.content.downloadUrl;
                        const tipusArxiu = file.content.fileType;

                        // Llegim el text de l'arxiu i el guardem a la variable
                        text = await getText(downloadUrl);

                        // Segons el tipus d'arxiu que sigui el preparem d'una manera o d'una altre
                        switch (tipusArxiu) {
                            case "txt":
                                text2 = await prepareTranscriptsTxt(text);
                                break;
                            case "vtt":
                                text2 = await prepareTranscriptsVtt(text);
                                break;
                            case "docx":
                                await context.sendActivity("No entenc el contingut d'aquest arxiu.");
                                await context.sendActivity("Siusplau passa'm un document .txt o .vtt.");
                                //text2 = await prepareTranscriptsDocx(text);
                                break;
                            default:
                                await context.sendActivity("No entenc el contingut d'aquest arxiu.");
                                await context.sendActivity("Siusplau passa'm un document .txt o .vtt.");
                                break;
                        }

                        // Aquí hauriem de contar els tokens, si supera el maxim avisar a l'usuari, sinó enviar la petició a la API
                        var tokens = openai_tokens.tokens(text2);
                        // El maxtokens son els tokens que ocupa la resposta, que estan incloso dins els 2048 tokens maxims
                        tokens = tokens - maxtokens; 

                        // Si els tokens superen 2048 avisem a l'usuari que el text es massa llarg, sinó li enviem el resum
                        if(tokens > 2048){
                            await context.sendActivity("Ocupa aquests tokens: " + tokens);
                            await context.sendActivity("El text es massa llarg, no podem fer-ne un resum.");
                        } else {
                            finalTranscription = await petititonOpenAiApi(text2);
                            await context.sendActivity("Ocupa aquests tokens: " + tokens);
                            await context.sendActivity("El resum de la reunió és: " + finalTranscription);
                        }
                    }
                }

            }

            // Pilla tots els membres de la conversa
            /*
            const teamDetails = await TeamsInfo.getMembers(context);
            if (teamDetails) {
                for (let cnt = 0; cnt < teamDetails.length; ++cnt) {
                    await context.sendActivity(`Teams Details Meeting Participants: ${ teamDetails[cnt].name }`);
                }
            } else {
                await context.sendActivity('This message did not come from a channel in a team.');
            }*/

            await next();
        });


        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;

            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    const holaPersona = `Benvingut/da ${ membersAdded[cnt].id}`;
                    await context.sendActivity(holaPersona);
                    await context.sendActivity("Si m'envies un fitxer amb les transcripcions d'una reunió t'envio el resum.");  
                }
            }
            await next();
        });

        // Nomes detecta quan afegim gent i marxa del chat
        this.onConversationUpdate(async (context, next) => {
            await context.sendActivity("onConversationUpdate"); 
            await next();
        });

        this.onEventActivity(async (context, next) => {
            await context.sendActivity("onEventActivity"); 
            await next();
        });

        this.onEvent(async (context, next) => {
            await context.sendActivity("onEvent"); 
            await next();
        });

        this.onInvokeActivity(async (context, next) => {
            await context.sendActivity("onInvokeActivity"); 
            await next();
        });

        this.onTurnActivity(async (context, next) => {
            await context.sendActivity("onTurnActivity"); 
            await next();
        });

        // Llegeix el text de la URL i el retorna
        async function getText(downloadUrl){
            var data = [];
            var text = null;
            return new Promise((resolve) => {
                https.get(downloadUrl, res => {
                    res.on('data', function(d) {
                        data.push(d);
                        var buffer = Buffer.concat(data);
                        text = buffer.toString('ascii');

                    }).on('end', function() {
                        resolve(text);
                    });
                });
            })
        }

        // Prepara el text de les transcripcions d'un fitxer txt
        async function prepareTranscriptsTxt(t){
            var textAux = "", myArr2 = "", textAux2 = "";

            myArr = t.split('-->');

            for(var i = 0; i < myArr.length; i++){

                myArr[i] = myArr[i].replace(regex, "");
                myArr[i] = myArr[i].replace(regex, "");  
                myArr2 = myArr[i].split('\n');

                textAux2 = "";

                for(var j = 0; j < myArr2.length; j++){
                    if(j == 1){
                        myArr2[j] = myArr2[j] + ": "; 
                    }
                    textAux2 = textAux2 + myArr2[j];
                }
                textAux = textAux + textAux2;
            }
            
            return petition + textAux + resum;
        }

        // Prepara el text de les transcripcions d'un fitxer vtt
        async function prepareTranscriptsVtt(t){
            var textAux = "";
            t = t.replace("WEBVTT", "");
            myArr = t.split('-->');
            for(var i = 0; i < myArr.length; i++){
                myArr[i] = myArr[i].replace(regex, '\n');
                myArr[i] = myArr[i].replace(regex, '');
                myArr[i] = myArr[i].replace("</v>", "");
                myArr[i] = myArr[i].replace("<v", "");
                myArr[i] = myArr[i].replace(">", ": ");
                textAux = textAux + myArr[i];
            }


            return petition + textAux + resum;
        }

        // Prepara el text de les transcripcions d'un fitxer docx
        // No es poden passar docx perquè no entenem el text
        async function prepareTranscriptsDocx(t){
            var textAux = "";
            myArr = t.split('-->');
            for(var i = 0; i < myArr.length; i++){
                myArr[i] = myArr[i].replace(regex, '\n');
                myArr[i] = myArr[i].replace(regex, '');
                textAux = textAux + myArr[i];
            }
            return petition + textAux + resum;
        }
        
        async function petititonOpenAiApi(transcription){

            finalTranscription = transcription;

            /*const gptResponse = await openai.complete({
              engine: 'davinci-instruct-beta',
              prompt: transcription,
              maxTokens: maxtokens,
              temperature: 0.7,
              topP: 1.0,
              presencePenalty: 0.0,
              frequencyPenalty: 0.0
            });

            finalTranscription = gptResponse.data.choices[0].text;*/
            return finalTranscription;
            
          };
    }
}

module.exports.MeetingBot = MeetingBot;
