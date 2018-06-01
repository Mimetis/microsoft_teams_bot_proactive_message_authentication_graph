import * as builder from "botbuilder";
import * as utils from "../utils";
import * as auth from "../authentication";
import { GraphServices } from "../services/graphServices";

interface PayloadAssistance {
  contactName: string;
  contactEmail: string;
  assistanceName: string;
  assistanceEmail: string;
  assistanceProblem: string;
}

function createLibrary(provider: auth.IOAuth2Provider): builder.Library {
  let lib = new builder.Library("newcomer");

  lib.dialog("welcome", [
    async (session, args) => {

      session.dialogData.payload = <PayloadAssistance>args;
      let pa: PayloadAssistance = args;

      session.send(`Wake up ! ... Néo ? ... Hum no... I guess ... Anyway, ${pa.assistanceName} needs your assistance. Here is the message he sent to us :`);
      session.send(pa.assistanceProblem);
      builder.Prompts.choice(
        session,
        `Do you want me to send him an email ?`,
        ["Yes", "No"]
      );
    },
    async (session, results, next) => {
      let resultLowerCase = results.response.entity.toLowerCase();

      if (resultLowerCase === "yes") {

        // Checking if the user is already connected, and if the token is not expired
        if (!auth.isAuthenticated(session, provider.providerName)) {
          session.beginDialog("auth:start");
        } else if (auth.isTokenExpired(session, provider.providerName)) {
          let token = auth.getUserToken(session, provider.providerName);
          token = await provider.getAccessTokenWithRefreshTokenAsync(token.refreshToken);
          auth.setUserToken(session, provider.providerName, token);
        }

        next();
      } else {
        return session.endDialog();
      }
    },
    async (session, results) => {
      builder.Prompts.text(
        session,
        `Please let me know what message to send to him on your behalf.`
      );
    },
    async (session, results) => {

      let messageBody = results.response.toLowerCase();

      // getting the token
      let userToken = auth.getUserToken(session, provider.providerName);

      let graphServices = new GraphServices();

      let pa: PayloadAssistance = session.dialogData.payload;

      session.sendTyping();

      // sending an email
      let emailSentResult = await graphServices.sendEmailAsync(pa.assistanceEmail, "Assistance from support", messageBody, userToken.accessToken);

      if (emailSentResult)
        session.send("Thanks! the email was correctly sended ! ");
      else
        session.send("Oops ... something did not worked correctly... ");

      session.endDialog();
    }
  ]);
  return lib.clone();
}


export default {
  createLibrary: (provider: auth.IOAuth2Provider) => createLibrary(provider)
}
