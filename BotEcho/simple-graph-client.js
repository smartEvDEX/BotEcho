var graph  = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

/**
 * This class is a wrapper for the Microsoft Graph API.
 * See: https://developer.microsoft.com/en-us/graph for more information.
 */
class SimpleGraphClient {
    constructor(token) {
        if (!token || !token.trim()) {
            throw new Error('SimpleGraphClient: Invalid token received.');
        }

        this._token = token;

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
        this.graphClient = graph.Client.init({
            authProvider: (done) => {
                done(null, this._token); // First parameter takes an error if you can't get an access token.
            }
        });
    }

    async getMe() {
        return await this.graphClient
            .api('/me')
            .get().then((res) => {
                return res;
            });
    }

    async getChannelReadBasic(team_id) {
        return await this.graphClient
        .api('/teams/'+ team_id +'/channels')
        .get().then((res) => {
            return res;
        });
    }

    async getChannelReadMessage(team_id, channel_id) {
        return await this.graphClient
        .api('/teams/' +  team_id +'/channels/' + channel_id + '/messages')
        .get().then((res) => {
            return res;
        });
    }


    // Quan acabi la reunió ha de fer aquesta petició
    async getCall(call_id) {
        return await this.graphClient
            .api('/communications/calls/'+ call_id)
            .get().then((res) => {
                return res;
            }).catch((res) => {
                return res;
            });
    }



    async getEvent() {
        return await this.graphClient
            .api('/me/calendar/events')
            .get().then((res) => {
                return res;
            }).catch((res) => {
                return res;
            });
    }


    async getCallRecord() {
        return await this.graphClient
            .api('/communications/callRecords/getDirectRoutingCalls(fromDateTime=2021-09-07,toDateTime=2021-09-08)')
            .get().then((res) => {
                return res;
            }).catch((res) => {
                return res;
            });
    }

    async getOnlineMeeting() {
        return await this.graphClient
            .api('/users/20b4df9467994f0bacbc74f68df41843/onlineMeetings/19:meeting_OTkxYTliNTctMjIzYS00YTc1LWFjYTQtNGRhNWM4NmVjOTE2@thread.v2')
            .get().then((res) => {
                return res;
            });
    }

    async getOnlineMeetingById() {
        return await this.graphClient
            .api('/communications/onlineMeetings/')
            .filter('VideoTeleconferenceId eq \'1211064580\'')
            .get().then((res) => {
                return res;
            });
    
    }

}

exports.SimpleGraphClient = SimpleGraphClient;