require("dotenv").config();
import * as restify from "restify"
import * as botbuilder from "botbuilder"
import * as teams from "botbuilder-teams"
import * as redis from "redis"

const cache = redis.createClient(6380, process.env.REDISCACHEHOSTNAME, {
    auth_pass: process.env.REDISCACHEKEY,
    tls: {
        servername: process.env.REDISCACHEHOSTNAME
    }
});

cache.on("error", (err) => {
    console.log(err);
});

const connector = new teams.TeamsChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});


// -- Bot definition -- //

const bot = new botbuilder.UniversalBot(connector);

bot.on("conversationUpdate", (activity) => {
    let conversationId = activity.address.conversation.id; // to be used in the fetchMembers request
    let serviceUrl = activity.address.serviceUrl; // to be used in the fetchMembers request
    let teamId = activity.sourceEvent.team.id; // to be used as key in Redis

    // save to Redis
    cache.set(teamId, JSON.stringify({ "conversationId": conversationId, "serviceUrl": serviceUrl }));
});

bot.dialog("/", (session) => {
    session.send("This the root dialog.");
});


// -- HTTP Server definition --- //

const server = restify.createServer();
server.use(restify.plugins.queryParser());
server.listen(process.env.port || process.env.PORT || 3979, () => {
    console.log(`${ server.name } listentning to ${ server.url }`);
});

server.post("/api/messages", connector.listen());

server.get("/api/users", async (req, res, next) => {
    // get team ID from query string
    if (req.query.teamId == undefined || req.query.teamId === "") {
        res.send(500, "Please provide the teamId parameter value.");
        return;
    }
    
    // ask Redis for channelId and serviceUrl for specified team
    const teamId: string = req.query.teamId;
    let cdata: ICacheData;
    try {
        cdata = await getCacheKeyAsync(teamId);

        // get user list
        connector.fetchMembers(cdata.serviceUrl, cdata.conversationId, (err, result) => {
            if (err) {
                console.log('error: ', err);
            } else {
                // return the list to client
                res.send(200, result, {"Content-Type": "application/json"});
            }
        });
    }
    catch (ex) {
        res.send("There was a problem while fetching members: " + ex)        
    }
});

// helper function to return data from cache using Promises
function getCacheKeyAsync(teamId: string): Promise<ICacheData> 
{
    return new Promise<ICacheData>((resolve, reject) => 
    {
        cache.get(teamId, function(err, result) {
            // something wrong happened with Redis
            if (err) {
                reject("Error getting data from Redis.");
            }
            // or redis succeeded, but key was not present
            else if (result == null) {
                reject("Key was not found in the cache.");
            }
            // or data retrieved successfuly
            else {
                const convData = JSON.parse(result);
                resolve(convData);
            }
        });
    });
}

interface ICacheData {
    conversationId: string;
    serviceUrl: string;
}