import * as builder from 'botbuilder';
import * as utils from '../utils';
import * as auth from '../authentication';

let uuidv4 = require("uuid/v4");

function createLibrary(oauthProvider: auth.IOAuth2Provider): builder.Library {

    let lib = new builder.Library('auth');

    // register the auth dialog in the library to be able to call it from a IDialogWaterfallStep element
    lib.dialog('login', new auth.AuthDialog(oauthProvider));

    lib.dialog('start', [
        async (session, args, next) => {

            let isAuthenticated = auth.isAuthenticated(session, oauthProvider.providerName);
            if (!isAuthenticated) {
                session.beginDialog("auth:login");
            } else if (auth.isTokenExpired(session, oauthProvider.providerName)) {
                let userToken = auth.getUserToken(session, oauthProvider.providerName);
                let newUserToken = await oauthProvider.getAccessTokenWithRefreshTokenAsync(userToken.refreshToken);
                auth.setUserToken(session, oauthProvider.providerName, newUserToken);

            }
            next();
        },
        async (session, results) => {
            return session.endDialog();
        }

    ]).triggerAction({ matches: /^auth/i, });;

    return lib.clone();
}

export default {
    createLibrary: (provider: auth.IOAuth2Provider) => createLibrary(provider)
}
