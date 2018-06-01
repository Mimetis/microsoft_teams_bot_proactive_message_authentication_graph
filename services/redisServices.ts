import * as redis from "redis";


/** 
 * structure of the data contained in the redis cache 
 */
export interface ITeamCacheData {
    conversationId?: string;
    teamsId?: string;
    tenantId?: string;
    serviceUrl?: string;
}

/**
 * redis services.
 */
export class RedisServices {

    private _key: string;
    private _cache: redis.RedisClient;
    private _teamCacheData: ITeamCacheData;


    get redisClient() {
        return this._cache;
    }

    constructor(settings: { hostName: string, password: string, key: string }) {

        // the key object in redis
        this._key = settings.key;


        // Redis config
        this._cache = redis.createClient(6380, settings.hostName,
            {
                auth_pass: settings.password,
                tls: {
                    servername: settings.hostName
                }
            });
    }




    /** 
     * Get or set a new redis cache for a team, storing the conversationid and the serviceurl 
     * */
    public async setTeamCacheAsync(data: ITeamCacheData): Promise<boolean> {
        return new Promise<boolean>((rs, rj) => {

            let dataString = JSON.stringify(data);

            this._cache.set(this._key, dataString, (err, result) => {
                if (err)
                    return rj(err);

                this._teamCacheData = data;
                return rs(true);
            });
        });
    }

    /**
     * Gets the conversation id and service url from a team id cache, stored in redis
     */
    public async getTeamCacheAsync(refresh: boolean = false): Promise<ITeamCacheData> {
        return new Promise<ITeamCacheData>((resolve, reject) => {

            if (this._teamCacheData != undefined && !refresh)
                return resolve(this._teamCacheData);

            this._cache.get(this._key, (err, result) => {

                if (err)
                    return reject(err);

                // or redis succeeded, but key was not present
                if (result == null)
                    return reject("Key was not found in the cache.");

                this._teamCacheData = JSON.parse(result);
                // or data retrieved successfuly
                resolve(this._teamCacheData);
            });
        });
    }
}




