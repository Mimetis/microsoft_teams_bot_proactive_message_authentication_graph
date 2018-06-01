import * as builder from "botbuilder";
import { AzureAdProvider, UserToken } from "./providers";
const randomNumber = require("random-number-csprng");


// How many digits the verification code should be
const verificationCodeLength = 6;

// How long the verification code is valid
const verificationCodeValidityInMilliseconds = 10 * 60 * 1000;       // 10 minutes

// Regexp to look for verification code in message
const verificationCodeRegExp = /\b\d{6}\b/;

// Gets the OAuth state for the given provider
export function getOAuthState(session: builder.Session, providerName: string): string {
    ensureProviderData(session, providerName);
    return (session.userData[providerName].oauthState);
}

// Sets the OAuth state for the given provider
export function setOAuthState(session: builder.Session, providerName: string, state: string): void {
    ensureProviderData(session, providerName);
    let data = session.userData[providerName];
    data.oauthState = state;
    session.save().sendBatch();
}

// Ensure that data bag for the given provider exists
export function ensureProviderData(session: builder.Session, providerName: string): void {
    if (!session.userData[providerName]) {
        session.userData[providerName] = {};
    }
}

/**
 *  Gets the validated user token for the given provider
 * */
export function getUserToken(session: builder.Session, providerName: string): UserToken {
    let token = getUserTokenUnsafe(session, providerName);
    return (token && token.verificationCodeValidated) ? token : null;
}

export function isAuthenticated(session: builder.Session, providerName: string): boolean {
    let token = getUserTokenUnsafe(session, providerName);
    return (token && token.verificationCodeValidated) ? true : false;
}

/** 
 * Check if the access token has expired 
 * */
export function isTokenExpired(session: builder.Session, providerName: string): boolean {
    let token = getUserTokenUnsafe(session, providerName);

    if (token.expirationTime < Date.now())
        return true;

    return false;
}

// Checks if the user has a token that is pending verification
export function isUserTokenPendingVerification(session: builder.Session, providerName: string): boolean {
    let token = getUserTokenUnsafe(session, providerName);
    return !!(token && !token.verificationCodeValidated && token.verificationCode);
}



// Sets the user token for the given provider
export function setUserToken(session: builder.Session, providerName: string, token: UserToken): void {
    ensureProviderData(session, providerName);
    let data = session.userData[providerName];
    data.userToken = token;
    session.save().sendBatch();
}

// Prepares a token for verification. The token is marked as unverified, and a new verification code is generated.
export async function prepareTokenForVerification(userToken: UserToken): Promise<void> {
    userToken.verificationCodeValidated = false;
    userToken.verificationCode = await generateVerificationCode();
    userToken.verificationCodeExpirationTime = Date.now() + verificationCodeValidityInMilliseconds;
}

// Finds a verification code in the text string
export function findVerificationCode(text: string): string {
    let match = verificationCodeRegExp.exec(text);
    return match && match[0];
}

// Validates the received verification code against what is expected
// If they match, the token is marked as validated and can be used by the bot. Otherwise, the token is removed.
export function validateVerificationCode(session: builder.Session, providerName: string, verificationCode: string): void {
    let tokenUnsafe = getUserTokenUnsafe(session, providerName);
    if (!tokenUnsafe.verificationCodeValidated) {
        if (verificationCode &&
            (tokenUnsafe.verificationCode === verificationCode) &&
            (tokenUnsafe.verificationCodeExpirationTime > Date.now())) {
            tokenUnsafe.verificationCodeValidated = true;
            setUserToken(session, providerName, tokenUnsafe);
        } else {
            console.warn("Verification code does not match.");

            // Clear out the token after the first failed attempt to validate
            // to avoid brute-forcing the verification code
            setUserToken(session, providerName, null);
        }
    } else {
        console.warn("Received unexpected login callback.");
    }
}

// Generate a verification code that the user has to enter to verify that the person that
// went through the authorization flow is the same one as the user in the chat.
async function generateVerificationCode(): Promise<string> {
    let verificationCode = await randomNumber(0, Math.pow(10, verificationCodeLength) - 1);
    return ("0".repeat(verificationCodeLength) + verificationCode).substr(-verificationCodeLength);
}

// Gets the user token for the given provider, even if it has not yet been validated
function getUserTokenUnsafe(session: builder.Session, providerName: string): UserToken {
    ensureProviderData(session, providerName);
    return (session.userData[providerName].userToken);
}


// Get locale from client info in event
export function getLocale(evt: builder.IEvent): string {
    let event = (evt as any);
    if (event.entities && event.entities.length) {
        let clientInfo = event.entities.find(e => e.type && e.type === "clientInfo");
        return clientInfo.locale;
    }
    return null;
}


// Load a Session corresponding to the given event
export function loadSessionAsync(bot: builder.UniversalBot, event: builder.IEvent): Promise<builder.Session> {
    return new Promise((resolve, reject) => {
        bot.loadSession(event.address, (err: any, session: builder.Session) => {
            if (err) {
                console.error("Failed to load session", { error: err, address: event.address });
                reject(err);
            } else if (!session) {
                console.error("Loaded null session", { address: event.address });
                reject(new Error("Failed to load session"));
            } else {
                let locale = getLocale(event);
                if (locale) {
                    (session as any)._locale = locale;
                    session.localizer.load(locale, (err2) => {
                        // Log errors but resolve session anyway
                        if (err2) {
                            console.error(`Failed to load localizer for ${locale}`, err2);
                        }
                        resolve(session);
                    });
                } else {
                    resolve(session);
                }
            }
        });
    });
}
