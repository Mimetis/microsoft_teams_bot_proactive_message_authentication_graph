import * as request from "request-promise";
import * as querystring from "querystring";
import * as authUtils from '../authUtils'
import { IOAuth2Provider, UserToken } from ".";

// =========================================================
// Azure Active Directory Endpoints
// =========================================================

const authorizationUrl = "https://login.microsoftonline.com/common/oauth2/authorize";
const accessTokenUrl = "https://login.microsoftonline.com/common/oauth2/token";
const callbackPath = "/auth/callback";

// Example implementation of AzureAD as an identity provider
// See https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-protocols-oauth-code
export class AzureAdProvider implements IOAuth2Provider {

    constructor(
        private clientId: string,
        private clientSecret: string
    ) {
    }

    get displayName(): string {
        return "Azure AD";
    }

    get providerName(): string {
        return "azuread";
    }


    // Return the url the user should navigate to to authenticate the app
    public getAuthorizationUrl(state: string, baseUrl: string, extraParams?: any): string {
        if (baseUrl.endsWith("/"))
            baseUrl = baseUrl.slice(0, baseUrl.length - 1);

        let params = {
            response_type: "code",
            response_mode: "query",
            client_id: this.clientId,
            redirect_uri: baseUrl + callbackPath,
            resource: "https://graph.microsoft.com",
            state: state,
        } as any;
        if (extraParams) {
            params = { ...extraParams, ...params };
        }
        return authorizationUrl + "?" + querystring.stringify(params);
    }

    // Redeem the authorization code for an access token
    public async getAccessTokenAsync(code: string, baseUrl: string): Promise<UserToken> {

        if (baseUrl.endsWith("/"))
            baseUrl = baseUrl.slice(0, baseUrl.length - 1);

        let params = {
            grant_type: "authorization_code",
            code: code,
            client_id: this.clientId,
            client_secret: this.clientSecret,
            redirect_uri: baseUrl + callbackPath,
            resource: "https://graph.microsoft.com",
        } as any;

        let responseBody = await request.post({ url: accessTokenUrl, form: params, json: true });
        return {
            accessToken: responseBody.access_token,
            expirationTime: responseBody.expires_on * 1000,
            refreshToken: responseBody.refresh_token
        };
    }

    public async getAccessTokenWithRefreshTokenAsync(refreshToken: string): Promise<UserToken> {

        let params = {
            grant_type: "refresh_token",
            refresh_token: refreshToken,
            client_id: this.clientId,
            client_secret: this.clientSecret
        } as any;

        let responseBody = await request.post({ url: accessTokenUrl, form: params, json: true });
        return {
            accessToken: responseBody.access_token,
            expirationTime: responseBody.expires_on * 1000,
            refreshToken: responseBody.refresh_token,
            verificationCodeValidated: true
        };
    }


}
