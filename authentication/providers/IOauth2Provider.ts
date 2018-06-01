
// =========================================================
// OAuth2 Provider
// =========================================================

// User token
export interface UserToken {
    // Access token
    accessToken: string;
    // Approximate expiration time of the access token, expressed as a number of milliseconds from midnight, January 1, 1970 Universal Coordinated Time (UTC)
    expirationTime: number;
    // Verification code
    verificationCode?: string;
    // Has the verification code been validated?
    verificationCodeValidated?: boolean;
    // Expiration time of verification code, expressed as a number of milliseconds from midnight, January 1, 1970 Universal Coordinated Time (UTC)
    verificationCodeExpirationTime?: number;
    // Refresh Token
    refreshToken: string
}

// Generic OAuth2 provider interface
export interface IOAuth2Provider {

    // Display name to use for the provider
    readonly displayName: string;

    // Display provider name for id purpose
    readonly providerName: string;

    // Return the url the user should navigate to to authenticate the app
    getAuthorizationUrl(state: string, baseUrl: string, extraParams?: any): string;

    // Redeem the authorization code for an access token
    getAccessTokenAsync(code: string, baseUrl: string): Promise<UserToken>;

    // Get a new accesstoken with a refresh token
    getAccessTokenWithRefreshTokenAsync(refreshToken: string): Promise<UserToken>;

}
