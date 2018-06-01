import * as builder from 'botbuilder';


function createLibrary(): builder.Library {
    let lib = new builder.Library('greeting');

    lib.dialog('start',
        async (session) => {
            session.send("Hey ! I'm your personal assistant. I don't do too many things for now, but let me some time, and my developer will add more and more features !!! .... I guess :)");
        });

    return lib.clone();
}

export default {
    createLibrary: () => createLibrary()
}

