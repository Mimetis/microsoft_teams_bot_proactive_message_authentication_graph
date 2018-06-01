import { URL } from "url";
import * as express from "express";
import * as builder from "botbuilder";
import * as teams from "botbuilder-teams";
import { RedisServices, ITeamCacheData } from "./redisServices";
import * as utils from "../utils";
import * as auth from "../authentication";

export class RouterServices {
  private _router = express.Router();
  private _connector: teams.TeamsChatConnector;

  constructor(
    private _bot: auth.AuthBot,
    private _cache: RedisServices,
    private _teamSettings: { appId: string; appName: string }
  ) {
    this._connector = utils.getConnector(_bot);
  }

  public listen() {
    return (req: any, res: any, next: any) => {
      // Save url in a static object
      utils.setSiteUrl(new URL(`https://${req.get('host')}`));
      // @ts-ignore
      this._connector.listen()(req, res, next);
    };
  }

  public routes(): express.Router {

    /** 
     * the unique route used by the bot
     */
    this._router.post('/api/messages', this.listen());


    /**
     * Handle callback from OAUTH2 authenticateion
     */
    this._router.get("/auth/callback", (req, res) => {
      this._bot.handleOAuthCallback(req, res);
    });

    /**
     * Send a proactive message to the bot, from the app
     */
    this._router.post("/api/welcome", async (req, res) => {

      // try to get the team cache
      let teamCache: ITeamCacheData;

      try {
        teamCache = await this._cache.getTeamCacheAsync();

      } catch (error) {
        res.status(400);
        return res.send({ error: "cache is not available" });
      }

      if (!teamCache ||
        !teamCache.serviceUrl ||
        !teamCache.teamsId ||
        !teamCache.tenantId ||
        !teamCache.conversationId
      ) {
        res.status(400);
        return res.send({ error: "cache is not available" });
      }

      if (!req.body) {
        res.status(400);
        return res.send({ error: "payload is not available" });
      }

      // getting the users
      let users = await utils.getMembersAsync(
        this._connector,
        teamCache.conversationId,
        teamCache.serviceUrl
      );

      // getting the right user if exists in the team
      let user = users.find(u => u.email === req.body.contactEmail || u.userPrincipalName == req.body.contactEmail);

      if (!user) {
        return res.sendStatus(404);
      }

      let settings = {
        serviceUrl: teamCache.serviceUrl,
        appId: this._teamSettings.appId,
        appName: this._teamSettings.appName,
        tenantId: teamCache.tenantId,
        userId: user.id
      };

      utils.sendProactiveMessage(this._bot, "newcomer:welcome", settings, req.body);

      return res.sendStatus(200);
    });



    return this._router;
  }
}
