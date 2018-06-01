import urlJoin = require("url-join");
import * as builder from "botbuilder";
import * as request from "request";
import * as teams from "botbuilder-teams";


// Helpers for working with messages

/** 
 * Creates a new Message
 * Unlike the botbuilder constructor, this defaults the textFormat to "xml"
 */
export function createMessage(session: builder.Session, text = "", textFormat = "xml"): builder.Message {
    return new builder.Message(session)
        .text(text)
        .textFormat("xml");
}

/** getting the TeamsChatConnector from the bot */
export function getConnector(_bot: builder.UniversalBot): teams.TeamsChatConnector {
    return <teams.TeamsChatConnector>_bot.connector('*');
}


/** Getting members from a teams conversation */
export async function getMembersAsync(connector: teams.TeamsChatConnector, conversationId: string, serviceUrl: string): Promise<Array<teams.ChannelAccount>> {
    return new Promise<Array<any>>((res, rej) => {
        connector.fetchMembers(serviceUrl, conversationId, (err, result) => {
            if (err)
                return rej(err);

            return res(result);
        });
    });
}

/** Getting team information */
export async function getTeamInfoAsync(connector: teams.TeamsChatConnector, teamId: string, serviceUrl: string): Promise<teams.TeamInfo> {
    return new Promise<teams.TeamInfo>((res, rej) => {
        connector.fetchTeamInfo(serviceUrl, teamId, (err, result) => {
            if (err)
                return rej(err);

            return res(result);
        });
    });
}


/** Getting the channels available from the team */
export async function getChannelsAsync(connector: teams.TeamsChatConnector, teamId: string, serviceUrl: string): Promise<teams.ChannelInfo[]> {

    return new Promise<any>((rs, rj) => {
        connector.fetchChannelList(serviceUrl, teamId, (err, result) => {
            if (err)
                return rj(err);

            return rs(result);
        });
    });
}


/** Geneate a new message and send it proactively to a specified user (1:1 conversation) */
export async function sendProactiveMessage(bot: builder.UniversalBot,
    dialogId: string,
    settings: { serviceUrl: string, appId: string, appName: string, tenantId: string, userId: string },
    args: any) {

    // generate the correct address
    var address = {
        channelId: "msteams",
        user: { id: settings.userId },
        channelData: {
            tenant: {
                id: settings.tenantId
            }
        },
        bot: {
            id: settings.appId,
            name: settings.appName
        },
        serviceUrl: settings.serviceUrl,
        useAuth: true
    };

    // begin a new dialog with the correct user
    bot.beginDialog(address, dialogId, args);

}

/** Get the channel id in the event */
export function getChannelId(event: builder.IEvent): string {
    let sourceEvent = event.sourceEvent;
    if (sourceEvent && sourceEvent.channel) {
        return sourceEvent.channel.id;
    }

    return "";
}

/** Get the team id in the event */
export function getTeamId(event: builder.IEvent): string {
    let sourceEvent = event.sourceEvent;
    if (sourceEvent && sourceEvent.team) {
        return sourceEvent.team.id;
    }
    return "";
}

/** Get the tenant id in the event */
export function getTenantId(event: builder.IEvent): string {
    let sourceEvent = event.sourceEvent;
    if (sourceEvent && sourceEvent.tenant) {
        return sourceEvent.tenant.id;
    }
    return "";
}

/** Returns true if this is message sent to a channel */
export function isChannelMessage(event: builder.IEvent): boolean {
    return !!getChannelId(event);
}

/** Returns true if this is message sent to a group (group chat or channel) **/
export function isGroupMessage(event: builder.IEvent): boolean {
    return event.address.conversation.isGroup || isChannelMessage(event);
}

/** Strip all mentions from text */
export function getTextWithoutMentions(message: builder.IMessage): string {
    let text = message.text;
    if (message.entities) {
        message.entities
            .filter(entity => entity.type === "mention")
            .forEach(entity => {
                text = text.replace(entity.text, "");
            });
        text = text.trim();
    }
    return text;
}

/** Get all user mentions */
export function getUserMentions(message: builder.IMessage): builder.IEntity[] {
    let entities = message.entities || [];
    let botMri = message.address.bot.id.toLowerCase();
    return entities.filter(entity => (entity.type === "mention") && (entity.mentioned.id.toLowerCase() !== botMri));
}

/** Create a mention entity for the user that sent this message */
export function createUserMention(message: builder.IMessage): builder.IEntity {
    let user = message.address.user;
    let text = "<at>" + user.name + "</at>";
    let entity = {
        type: "mention",
        mentioned: user,
        entity: text,
        text: text,
    };
    return entity;
}

/** Gets the serviceUrl from a message */
export function getServiceUrl(message: builder.IMessage): string {
    return (<builder.IChatConnectorAddress>message.address).serviceUrl;
}
