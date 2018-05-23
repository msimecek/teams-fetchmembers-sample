# Microsoft Teams fetchMembers() sample - Node.js

intro here :)

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