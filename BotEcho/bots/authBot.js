const { DialogBot } = require('./dialogBot');

class AuthBot extends DialogBot {
    /**
     *
     * @param {ConversationState} conversationState
     * @param {UserState} userState
     * @param {Dialog} dialog
     */
    constructor(conversationState, userState, dialog) {
        super(conversationState, userState, dialog);

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    //await context.sendActivity('Welcome to AuthenticationBot. Type anything to get logged in. Type \'logout\' to sign-out.');
                }
            }

            await next();
        });

        this.onTokenResponseEvent(async (context, next) => {
            await context.sendActivity('Running dialog with Token Response Event Activity.');
            await this.dialog.run(context, this.dialogState);

            await next();
        });
    }

    async handleTeamsSigninVerifyState(context, state) {
        await this.dialog.run(context, this.dialogState);
    }
}

module.exports.AuthBot = AuthBot;