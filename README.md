# Microsoft Teams fetchMembers() sample - Node.js

This sample demonstrates how to get on-demand list of members of a particular team in Microsoft Teams. In this scenario the team owner is sideloading an app to Teams in order to use a *configurable tab*. When the tab is loaded, it needs to get a list of members for further processing. This task can be generalized to an external API call.

1. A *chatbot* exposes two HTTP endpoints: `(POST) /api/messages` for incoming messages and `(GET) /api/users` for getting the on-demand user list.
1. When application package is sideloaded, *chatbot* gets the `conversationUpdate` message.
1. It extracts `teamId`, `conversationId` and `serviceUrl` from this message.
1. It stores these values in *Azure Redis cache* with `teamId` as key.
1. When *external resource* (such as the tab) wants to get the user list, it calls the `/api/users` endpoint with `teamId` as URL parameter.
1. *Chatbot* asks Redis cache for `conversationId` and `serviceUrl`.
1. And then calls `fechMembers()` with these values.
1. Team roster is returned as JSON.

See the [Microsoft Teams documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/bots/bots-context#fetching-the-team-roster) for more details on `fetchMembers`.

## Prerequisites

To use this sample you need to provision several resources in [Microsoft Azure](https://azure.microsoft.com) first:

1. [Bot Channels Registration](https://docs.microsoft.com/en-us/azure/bot-service/bot-service-quickstart-registration?view=azure-bot-service-3.0) - make note of the **application ID** & **key**.
1. [Redis Cache](https://docs.microsoft.com/en-us/azure/redis-cache/cache-nodejs-get-started)  make note of the **key** & **hostname**.
1. [TypeScript](https://www.typescriptlang.org/#download-links)

You will also need a [Microsoft Teams](https://teams.microsoft.com) team and permissions to [sideload packages](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload) to this team.

## Quick start

Start by cloning this repo:

```cmd
git clone https://github.com/msimecek/teams-fetchmembers-sample.git
```

Rename `.env.template` to `.env` and fill your bot's **application ID** & **key** and your **Redis key** and **hostname**.

```
MicrosoftAppId=appidappidappidappid
MicrosoftAppPassword=secretsecretsecret

REDISCACHEHOSTNAME=mycache.redis.cache.windows.net
REDISCACHEKEY=keykeykeykey
```

Restore packages, compile TypeScript and start the app:

```bat
cd teams-fetchmembers-sample
npm install
tsc
node dist/app.js
```

Start [Ngrok](https://ngrok.com/) on port **3979** and copy the **https URL**.

```
ngrok http 3979 --host-header="localhost:3979"
```

Set your bot's messaging URL in Azure to your Ngrok URL with `/api/messages` attached.

```
# example:
https://abcdabcd.ngrok.io/api/messages
```

Create Teams app package for your bot and sideload it to your team. It's described [in the documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/bots/bots-test#adding-a-bot-to-a-team-for-use-in-channels).

Open a web browser and call the `/api/users` endpoint of your bot.

## Code walkthrough

// TODO :)