import { URL } from "url";

let _siteUrl: URL;
export function setSiteUrl(value: URL) {
    this._siteUrl = value;
}

export function getSiteUrl(): URL {
    return this._siteUrl;
}
