import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as auth from ".";
import * as utils from "../utils";
import { Request, Response } from "express";

// =========================================================
// Auth Bot
// =========================================================

export class AuthBot extends builder.UniversalBot {

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any
    ) {
        super(_connector, botSettings);
        this.set("persistConversationData", true);

        // Handle generic invokes
        let teamsConnector = this._connector as msteams.TeamsChatConnector;
        teamsConnector.onInvoke(async (event, cb) => {
            try {
                await this.onInvoke(event, cb);
            } catch (e) {
                console.error("Invoke handler failed", e);
                cb(e, null, 500);
            }
        });
        teamsConnector.onSigninStateVerification(async (event, query, cb) => {
            try {
                await this.onInvoke(event, cb);
            } catch (e) {
                console.error("Signin state verification handler failed", e);
                cb(e, null, 500);
            }
        });

    }

    // Handle OAuth callbacks
    // The provider name is in the route, which is defined as "/auth/:provider/callback"
    public async handleOAuthCallback(req: Request, res: Response): Promise<void> {
        const stateString = req.query.state as string;
        const state = JSON.parse(stateString);
        const authCode = req.query.code;
        const provider: auth.IOAuth2Provider = this.botSettings.oauthProvider;
        let verificationCode = "";

        // Load the session from the address information in the OAuth state.
        // We'll later validate the state to check that it was not forged.
        let session: builder.Session;
        let address: builder.IAddress;
        try {
            address = state.address as builder.IAddress;
            session = await auth.loadSessionAsync(this, {
                type: "invoke",
                agent: "botbuilder",
                source: address.channelId,
                sourceEvent: {},
                address: address,
                user: address.user,
            });
        } catch (e) {
            console.warn("Failed to get address from OAuth state", e);
        }

        if (session &&
            (auth.getOAuthState(session, provider.providerName) === stateString) &&     // OAuth state matches what we expect
            authCode) {                                                         // User granted authorization
            try {
                // Redeem the authorization code for an access token, and store it provisionally
                // The bot will refuse to use the token until we validate that the user in the chat
                // is the same as the user who went through the authorization flow, using a verification code
                // that needs to be presented by the user in the chat.

                let userToken = await provider.getAccessTokenAsync(authCode, utils.getSiteUrl().origin);

                await auth.prepareTokenForVerification(userToken);
                auth.setUserToken(session, provider.providerName, userToken);

                verificationCode = userToken.verificationCode;
            } catch (e) {
                console.error("Failed to redeem code for an access token", e);
            }
        } else {
            console.warn("State does not match expected state parameter, or user denied authorization");
        }

        // Render the page shown to the user
        if (verificationCode) {
            // If we have a verification code, we were able to redeem the code successfully. Render a page
            // that calls notifySuccess() with the verification code, or instructs the user to enter it in chat.
            res.render("oauth-callback-success", {
                verificationCode: verificationCode,
                providerName: provider.displayName,
            });

            // The auth flow resumes when we receive the verification code response, which can happen either:
            // 1) through notifySuccess(), which is handled in BaseIdentityDialog.handleLoginCallback()
            // 2) by user entering it in chat, which is handled in BaseIdentityDialog.onMessageReceived()

        } else {
            // Otherwise render an error page
            res.render("oauth-callback-error", {
                providerName: provider.displayName,
            });
        }
    }

    // Handle incoming invoke
    private async onInvoke(event: builder.IEvent, cb: (err: Error, body: any, status?: number) => void): Promise<void> {
        let session = await auth.loadSessionAsync(this, event);
        if (session) {
            // Invokes don't participate in middleware

            // Simulate a normal message and route it, but remember the original invoke message
            let payload = (event as any).value;
            let fakeMessage: any = {
                ...event,
                text: payload.command + " " + JSON.stringify(payload),
                originalInvoke: event,
            };

            session.message = fakeMessage;
            session.dispatch(session.sessionState, session.message, () => {
                session.routeToActiveDialog();
            });
        }
        cb(null, "");
    }
}
