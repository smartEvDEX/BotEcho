
const { SimpleGraphClient } = require('./simple-graph-client');

/**
 * These methods call the Microsoft Graph API. The following OAuth scopes are used:
 * 'openid' 'profile' 'User.Read'
 * for more information about scopes see:
 * https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference
 */
class OAuthHelpers {
    /**
     * Send the user their Graph Display Name from the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     */
    static async listMe(context, tokenResponse) {
        await context.sendActivity("Entra al listMe");
        if (!context) {
            throw new Error('OAuthHelpers.listMe(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.listMe(): `tokenResponse` cannot be undefined.');
        }

        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();

        await context.sendActivity(`You are ${ me.displayName }.`);
    }

    /**
     * Send the user their Graph Email Address from the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     */
    static async listEmailAddress(context, tokenResponse) {
        if (!context) {
            throw new Error('OAuthHelpers.listEmailAddress(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.listEmailAddress(): `tokenResponse` cannot be undefined.');
        }

        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();

        await context.sendActivity(`Your email: ${ me.mail }.`);
    }



    static async listTranscriptions(context, tokenResponse, id) {
        console.log(context);
        if (!context) {
            throw new Error('OAuthHelpers.listTranscriptions(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.listTranscriptions(): `tokenResponse` cannot be undefined.');
        }

        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);

        const info = await client.getCall(id);
        //Hem de tractar el que ens importi
        await context.sendActivity(info);
    }

    static async listEvents(context, tokenResponse) {
        console.log(context);
        if (!context) {
            throw new Error('OAuthHelpers.listTranscriptions(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.listTranscriptions(): `tokenResponse` cannot be undefined.');
        }

        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);

        const info = await client.getOnlineMeetingById();
        //Hem de tractar el que ens importi
        await context.sendActivity(info);
    }

    static async listCallRecords(context, tokenResponse) {
        console.log(context);
        if (!context) {
            throw new Error('OAuthHelpers.listTranscriptions(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.listTranscriptions(): `tokenResponse` cannot be undefined.');
        }

        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);

        const info = await client.getOnlineMeeting();
        //Hem de tractar el que ens importi
        await context.sendActivity(info);
    }
}

exports.OAuthHelpers = OAuthHelpers;