import * as request from "request-promise";
import * as auth from "../authentication"
import { Client } from "@microsoft/microsoft-graph-client";
import { Message } from "@microsoft/microsoft-graph-types";

const graphProfileUrl = "https://graph.microsoft.com/v1.0/me";


export class GraphServices {

    async sendEmailAsync(to: string, subject: string, content: string, accessToken: string): Promise<boolean> {

        return new Promise<boolean>((rs, rj) => {
            var client = Client.init({
                authProvider: async (done) => {
                    done(null, accessToken); //first parameter takes an error if you can't get an access token
                }
            });

            // construct the email object
            const mail = {
                subject: subject,
                toRecipients: [{
                    emailAddress: { address: to }
                }],
                body: {
                    content: content,
                    contentType: "html"
                }
            }

            client
                .api('/users/me/sendMail')
                .post({ message: mail }, (err, res) => {

                    if (err)
                        return rs(false);

                    return rs(true);
                })
        });

    }


}