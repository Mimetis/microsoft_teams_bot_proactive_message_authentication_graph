import * as builder from "botbuilder";
import * as authUtils from "./authUtils";
import * as utils from "../utils";
import { IOAuth2Provider, UserToken } from "./providers";
import * as request from "request-promise";
let uuidv4 = require("uuid/v4");

const graphProfileUrl = "https://graph.microsoft.com/v1.0/me";

// Base identity dialog
export class AuthDialog extends builder.IntentDialog {


    constructor(private authProvider: IOAuth2Provider) {
        super();

        this.onBegin((session, args, next) => { this.onDialogBegin(session, args, next); });
        this.onDefault((session) => { this.onMessageReceived(session); });

        this.matches(/SignIn/, (session) => { this.handleLogin(session, utils.getSiteUrl().origin); });
        this.matches(/ShowProfile/, async (session) => { await this.showUserProfile(session); });
        this.matches(/SignOut/, (session) => { this.handleLogout(session); });
        this.matches(/Back/, (session) => { session.endDialog(); });
    }

    // Get an authorization url for the identity provider
    // This allows derived dialogs to pass additional parameters to their identity provider implementation
    // by overriding this method. See the implementation of AzureADv1Dialog.
    protected getAuthorizationUrl(session: builder.Session, state: string): string {
        return this.authProvider.getAuthorizationUrl(state, utils.getSiteUrl().origin);
    }

    // Handle start of dialog
    private async onDialogBegin(session: builder.Session, args: any, next: () => void): Promise<void> {
        session.dialogData.isFirstTurn = true;

        // try to get the user token, to check auth
        let userToken = authUtils.getUserToken(session, this.authProvider.providerName);

        if (!userToken || !userToken.verificationCodeValidated) {
            await this.handleLogin(session, utils.getSiteUrl().origin);
        }
        else {
            await this.showUserProfile(session);
            session.endDialog();
        }
        next();
    }

    // Handle message
    private async onMessageReceived(session: builder.Session): Promise<void> {
        let messageAsAny = session.message as any;

        // for future purpose. let's check if we are in a 1:1 conversation, or not
        // let isChannelMessage = (session.message.sourceEvent && session.message.sourceEvent.channel && session.message.sourceEvent.channel.id);
        // console.log("message is in a channel ? : " + isChannelMessage);

        if (messageAsAny.originalInvoke) {
            // This was originally an invoke message, see if it is signin/verifyState
            let event = messageAsAny.originalInvoke;

            if (event.name === "signin/verifyState") {
                await this.handleLoginCallback(session);
            } else {
                console.warn(`Received unrecognized invoke "${event.name}"`);
            }
        } else {
            // See if we are waiting for a verification code and got one
            if (authUtils.isUserTokenPendingVerification(session, this.authProvider.providerName)) {
                let verificationCode = authUtils.findVerificationCode(session.message.text);
                authUtils.validateVerificationCode(session, this.authProvider.providerName, verificationCode);

                // End of auth flow: if the token is marked as validated, then the user is logged in

                if (authUtils.getUserToken(session, this.authProvider.providerName)) {
                    await this.showUserProfile(session);
                    session.endDialog();

                } else {
                    session.send(`Sorry, there was an error signing in to ${this.authProvider.displayName}. Please try again.`);
                }

            } else {
                // Unrecognized input
                session.send("I didn't understand. Exit dialog.");
                session.endDialog();
            }
        }
    }




    // Handle user login callback
    private async handleLoginCallback(session: builder.Session): Promise<void> {
        let messageAsAny = session.message as any;
        let verificationCode = messageAsAny.originalInvoke.value.state;

        authUtils.validateVerificationCode(session, this.authProvider.providerName, verificationCode);

        // End of auth flow: if the token is marked as validated, then the user is logged in

        if (authUtils.getUserToken(session, this.authProvider.providerName)) {
            await this.showUserProfile(session);
            session.endDialog();
        } else {
            session.send(`Sorry, there was an error signing in to ${this.authProvider.displayName}. Please try again.`);
        }
    }

    // Handle user logout request
    private async handleLogout(session: builder.Session): Promise<void> {
        if (!authUtils.getUserToken(session, this.authProvider.providerName)) {
            session.send(`You're already signed out of ${this.authProvider.displayName}.`);
        } else {
            authUtils.setUserToken(session, this.authProvider.providerName, null);
            session.send(`You're now signed out of ${this.authProvider.displayName}.`);
        }

        session.endDialog();
    }

    // Handle user login request
    private async handleLogin(session: builder.Session, baseUrl: string): Promise<void> {

        if (baseUrl.endsWith("/"))
            baseUrl = baseUrl.slice(0, baseUrl.length - 1);

        if (authUtils.getUserToken(session, this.authProvider.providerName)) {
            session.send(`You're already signed in to ${this.authProvider.displayName}.`);

            session.endDialog();
        } else {
            // Create the OAuth state, including a random anti-forgery state token
            let address = session.message.address;
            let state = JSON.stringify({
                securityToken: uuidv4(),
                address: {
                    user: { id: address.user.id, },
                    conversation: { id: address.conversation.id },
                },
            });
            authUtils.setOAuthState(session, this.authProvider.providerName, state);

            // Create the authorization URL
            let authUrl = this.getAuthorizationUrl(session, state);

            // Build the sign-in url
            let signinUrl = baseUrl + `/html/auth-start.html?authorizationUrl=${encodeURIComponent(authUrl)}`;

            // The fallbackUrl specifies the page to be opened on mobile, until they support automatically passing the
            // verification code via notifySuccess(). If you want to support only this protocol, then you can give the
            // URL of an error page that directs the user to sign in using the desktop app. The flow demonstrated here
            // gracefully falls back to asking the user to enter the verification code manually, so we use the same
            // signin URL as the fallback URL.
            let signinUrlWithFallback = signinUrl + `&fallbackUrl=${encodeURIComponent(signinUrl)}`;

            // Send card with signin action
            let msg = new builder.Message(session)
                .addAttachment(new builder.HeroCard(session)
                    .text(`To be able to send an email (on your behalf), I need your consent. Can you please login ?`)
                    .buttons([
                        new builder.CardAction(session)
                            .type("signin")
                            .value(signinUrlWithFallback)
                            .title("Yes"),
                        builder.CardAction.messageBack(session, "{}", "No")
                            .text("Back")
                            .displayText("No"),
                    ]));


            session.send(msg);

            // The auth flow resumes when we handle the identity provider's OAuth callback in AuthBot.handleOAuthCallback()
        }
    }

    // Show user profile
    public async showUserProfile(session: builder.Session): Promise<void> {


        let userToken = authUtils.getUserToken(session, this.authProvider.providerName);

        let profile: any = {};

        if (userToken) {
            try {

                let options = {
                    url: graphProfileUrl,
                    json: true,
                    headers: {
                        "Authorization": `Bearer ${userToken.accessToken}`,
                    },
                };
                return await request.get(options);

            } catch (error) {
                console.error(error);
            }
        }

        if (!userToken) {
            session.send("Please sign in to AzureAD so I can access your profile.");
        } else if (!profile) {
            session.send("Your consent seems to be unavailable, please sign in again.");
        } else {
            let profileCard = new builder.ThumbnailCard()
                .title(profile.displayName)
                .subtitle(profile.mail)
                .text(`${profile.jobTitle}<br/> ${profile.officeLocation}`);
            session.send(new builder.Message().addAttachment(profileCard));

        }

    }
}
