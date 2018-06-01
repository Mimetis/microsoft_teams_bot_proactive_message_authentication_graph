import * as builder from 'botbuilder';

export function createLibrary() {

    let lib = new builder.Library('reset');

    lib.dialog('everything', [
        (session, args, skip) => {
            Object.keys(session.userData).forEach(key => delete session.userData[key]);
            session.endConversation();
            session.send('Conversation ended and restarted, profile deleted.');
            session.endDialog();
        }
    ]).triggerAction({ matches: /^reset/i, });

    lib.dialog('conversation', [
        (session, args, skip) => {
            session.endConversation();
            session.send('Conversation ended and restarted.');
            session.endDialog();
        }
    ]).triggerAction({ matches: /^restart/i, });

    return lib.clone();
}

export default {
    createLibrary: () => createLibrary()
}
