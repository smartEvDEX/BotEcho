// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.


// https://www.npmjs.com/package/@microsoft/microsoft-graph-client#3-make-requests-to-the-graph

const { 
    ActivityHandler,
    TurnContext, 
    MessageFactory,
    TeamsInfo
} = require('botbuilder');

const OpenAI = require('openai-api');
const OpenAITokens = require('openai-nodejs');
const config = require('./config.json');
const https = require('https');
const path = require('path');


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


class MeetingBot extends ActivityHandler {
    constructor() {
        super();

        // When the bot recive a message
        this.onMessage(async (context, next) => {

            const attachments = context.activity.attachments;
            TurnContext.removeRecipientMention(context.activity);

            var text;
            if(context.activity.text != null){
                text = context.activity.text.trim().toLocaleLowerCase();
            } else {
                text = ""; 
            }
            
            // Detect if there is a attached file
            if (attachments == undefined) {

                // If there isn't any attached files the bot inform what do it want
                if (text.includes('hello')) {
                    await context.sendActivity("Hi!"); 
                    await context.sendActivity("If you send me a file with the transcripts of a meeting I will send you the summary.");
                    await context.sendActivity("Remember that I only accept .txt and .vtt files.");
                } else if (text.includes('thank') || text.includes('thanks')) {
                    await context.sendActivity("You're welcome.");
                    await context.sendActivity("If you need any other transcriptions send me the file.");
                } else if (text.includes('bye')) {
                    await context.sendActivity("If you need any other transcriptions send me the file.");
                    await context.sendActivity("Bye!");
                } else {
                    await context.sendActivity("Hello!"); 
                    await context.sendActivity("If you send me a file with the transcripts of a meeting I will send you the summary.");
                    await context.sendActivity("Remember that I only accept .txt and .vtt files.");
                }

            } else {

                // For all the attached files
                for (let i = 0; i < attachments.length; i++){
                    const file = attachments[i];
                    // In case the ot has mention in a group conversation
                    if(file.contentType == "text/html"){

                        if (text.includes('hello')) {
                            await context.sendActivity("Hi!"); 
                            await context.sendActivity("If you send me a file with the transcripts of a meeting I will send you the summary.");
                            await context.sendActivity("Remember that I only accept .txt and .vtt files.");
                        } else if (text.includes('thank') || text.includes('thanks')) {
                            await context.sendActivity("You're welcome.");
                            await context.sendActivity("If you need any other transcriptions send me the file.");
                        } else if (text.includes('bye')) {
                            await context.sendActivity("If you need any other transcriptions send me the file.");
                            await context.sendActivity("Bye!");
                        } else {
                            await context.sendActivity("Welcome!"); 
                            await context.sendActivity("If you send me a private message with the transcriptions file of a meeting I will send you the summary.");
                            await context.sendActivity("Remember that I only accept .txt and .vtt files."); 
                        }

                    } else {
                        // If has a file in a une vs one conversation
                        transcription = null;
                        finalTranscription = null;
                        text = null;
                        text2 = null;
            
                        // We have to save the download URL of the file and the type of file
                        const downloadUrl = file.content.downloadUrl;
                        const tipusArxiu = file.content.fileType;

                        await context.sendActivity("Download URL: " + downloadUrl);

                        // We read the text from the file and save it to the variable text
                        text = await getText(downloadUrl);

                        // Depending on the type of file, we prepare it in one way or another
                        // THe bot accept files of vtt type or txt type
                        switch (tipusArxiu) {
                            case "txt":
                                text2 = await prepareTranscriptsTxt(text);
                                break;
                            case "vtt":
                                text2 = await prepareTranscriptsVtt(text);
                                break;
                            case "docx":
                                await context.sendActivity("I do not understand the contents of this file.");
                                await context.sendActivity("Please send me a document of .txt or .vtt type.");
                                text2 = null;
                                //text2 = await prepareTranscriptsDocx(text);
                                break;
                            default:
                                await context.sendActivity("I do not understand the contents of this file.");
                                await context.sendActivity("Please send me a document of .txt or .vtt type.");
                                text2 = null;
                                break;
                        }

                        // In case that we have read the file
                        if( text2 != null){
                            // We count the amount of tokens the text have
                            var tokens = openai_tokens.tokens(text2);
                            // The maxtokens are the tokens that occupy the answer, which are included in the maximum 2048 tokens
                            tokens = tokens - maxtokens; 
    
                            // If the tokens exceed 2048 we warn the user that the text is too long, otherwise we send the summary
                            if(tokens > 2048){
                                //await context.sendActivity("Ocupa aquests tokens: " + tokens);
                                await context.sendActivity("The text is too long, we can't summarize it.");
                            } else {
                                finalTranscription = await petititonOpenAiApi(text2);
                                //await context.sendActivity("Ocupa aquests tokens: " + tokens);
                                await context.sendActivity("The summary of the meeting is: \n" + finalTranscription);
                            }
                        }
                    }
                }

            }
            await next();
        });

        // When a member is added in the conversaion
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;

            // The bot explain what it does
            await context.sendActivity("Welcome!"); 
            await context.sendActivity("If you send me a file with the transcripts of a meeting I will send you the summary.");
            await context.sendActivity("Remember that I only accept .txt and .vtt files.");

            await next();
        });

        // It only detects when we add people and leaves the chat
        this.onConversationUpdate(async (context, next) => {
            //await context.sendActivity("onConversationUpdate"); 
            await next();
        });

        

        // Read the URL text and return it
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

        // Prepares the text of the transcripts of a txt file
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

        // Prepares the text of the transcripts of a vtt file
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

        // Prepares the text of the transcripts of a docx file
        // Docx cannot be passed because we do not understand the text
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
        

        // Function that makes the request to the OpenAI API and returns the response 
        async function petititonOpenAiApi(transcription){

            finalTranscription = transcription;

            const gptResponse = await openai.complete({
              engine: 'davinci-instruct-beta',
              prompt: transcription,
              maxTokens: maxtokens,
              temperature: 0.7,
              topP: 1.0,
              presencePenalty: 0.0,
              frequencyPenalty: 0.0
            });

            finalTranscription = gptResponse.data.choices[0].text;
            return finalTranscription;
            
          };
    }
}

module.exports.MeetingBot = MeetingBot;
