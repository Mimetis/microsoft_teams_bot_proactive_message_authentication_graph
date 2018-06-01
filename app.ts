import { URL } from "url";
import * as express from "express";
import * as path from "path";
import * as favicon from "serve-favicon";
import * as logger from "morgan";
import * as cookieParser from "cookie-parser";
import { json, urlencoded } from "body-parser";
import * as builder from "botbuilder";
import * as teams from "botbuilder-teams";
import * as dotenv from 'dotenv';
import { RedisServices, ITeamCacheData } from "./services/redisServices";
import { RouterServices } from "./services/routerServices";
import * as libs from "./librairies";
import * as utils from "./utils";
import { RedisStorage } from "./storage/redisStorage";
import * as auth from "./authentication";
import * as exphbs from "express-handlebars";

var app = express();

// loading the env variables
const result = dotenv.config();

// throw an error if no .env is present
if (result.error) {
  throw result.error;
}

// getting env variables
const teamsSettings = {
  appId: process.env.TEAMS_APP_ID,
  appPassword: process.env.TEAMS_APP_PASSWORD,
  appName: process.env.TEAMS_APP_NAME
}
const redisSettings = {
  hostName: process.env.REDIS_CACHE_HOSTNAME,
  password: process.env.REDIS_CACHE_PASSWORD,
  key: process.env.REDIS_CACHE_TEAM_KEY
}

app.use(express.static(path.join(__dirname, "../public")));
app.use(logger('dev'));
app.use(json());
app.use(urlencoded({ extended: false }));
app.use(cookieParser());

let handlebars = exphbs.create({
  extname: ".hbs",
  helpers: {
    appId: () => { return teamsSettings.appId; },
  },
});
app.engine("hbs", handlebars.engine);
app.set("view engine", "hbs");

// Creating the connector espcially for Teams
let connector = new teams.TeamsChatConnector(teamsSettings);

// creating the redis cache
let cache = new RedisServices(redisSettings);

// set bot storage
let botStorage = new RedisStorage(cache.redisClient);

// create the oauth2 provider
let oauthProvider = new auth.AzureAdProvider(teamsSettings.appId, teamsSettings.appPassword);

let botSettings = {
  storage: botStorage,
  oauthProvider: oauthProvider,
  defaultDialogId: "greeting:start"
};

let bot = new auth.AuthBot(connector, botSettings);

// Log bot errors 
bot.on("error", (error: Error) => {
  console.error(error.message, error);
});

// Create routes
let rs = new RouterServices(bot, cache, teamsSettings);

// use those routes
app.use('/', rs.routes());


// Adding libraries 
bot.library(libs.greet.createLibrary());
bot.library(libs.reset.createLibrary());
bot.library(libs.auth.createLibrary(oauthProvider));
bot.library(libs.newcomer.createLibrary(oauthProvider));


// send greetings to user when joining the conversation
bot.on('conversationUpdate', async (message) => {

  if (message.membersAdded) {
    message.membersAdded.forEach((identity: builder.IIdentity) => {

      // bot is just registered (application is just installed)
      if (identity.id === message.address.bot.id) {
        setredisCache(message)
      }

      // a new member is coming
      if (identity.id !== message.address.bot.id) {
        bot.beginDialog(message.address, "greeting:start");
      }
    });
  }
});


async function setredisCache(message: any) {

  let teamCache: ITeamCacheData = {};

  if (message.address && message.address.conversation)
    teamCache.conversationId = message.address.conversation.id;

  if (message.address && message.address.serviceUrl)
    teamCache.serviceUrl = message.address.serviceUrl;

  teamCache.teamsId = utils.getTeamId(message);
  teamCache.tenantId = utils.getTenantId(message);

  if (teamCache.serviceUrl && teamCache.teamsId && teamCache.tenantId && teamCache.conversationId) {
    await cache.setTeamCacheAsync(teamCache);
  }

}

// regex triggers
// bot.beginDialogAction('restart', 'reset:conversation', { matches: new RegExp("^<at>" + teamsSettings.appName + "</at> restart|^restart/i") });
// bot.beginDialogAction('reset', 'reset:everything', { matches: new RegExp("^<at>" + teamsSettings.appName + "</at> reset|^reset/i") });
// bot.beginDialogAction('auth', 'auth:start', { matches: new RegExp("^<at>" + teamsSettings.appName + "</at> auth|^auth/i") });

// bot.beginDialogAction('restart', 'reset:conversation', { matches: /^restart/i });
// bot.beginDialogAction('reset', 'reset:everything', { matches: /^reset/i });
// bot.beginDialogAction('auth', 'auth:start', { matches: /^auth/i });

export default app;
