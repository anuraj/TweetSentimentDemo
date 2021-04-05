process.env["NODE_TLS_REJECT_UNAUTHORIZED"] = 0;
const url = 'YOUR_TEAMS_CHANNEL_WEBHOOK_URL';
const Twit = require('twit');
const express = require('express');
const app = express();
const server = require('http').Server(app);
const { TextAnalyticsClient, AzureKeyCredential } = require("@azure/ai-text-analytics");
const CosmosClient = require("@azure/cosmos").CosmosClient;
const { IncomingWebhook } = require('ms-teams-webhook');
var io = require('socket.io')(server);

io.on('connection', function (socket) {
    console.log("Client connected");
});

var T = new Twit({
    consumer_key: 'YOUR_CONSUMER_KEY',
    consumer_secret: 'YOUR_CONSUMER_SECRET',
    access_token: 'YOUR_ACCESS_TOKEN',
    access_token_secret: 'YOUR_ACCESS_TOKEN_SECRET',
    timeout_ms: 60 * 1000
});

app.use(express.static('public'))

const client = new TextAnalyticsClient("YOUR_TEXT_ANALYTICS_ENDPOINT", new AzureKeyCredential("YOUR_TEXT_ANALYTICS_KEY"));
const cosmosClient = new CosmosClient({ endpoint: "YOUR_COSMOS_DB_ENDPOINT", key: "YOUR_COSMOS_DB_KEY" });
const database = cosmosClient.database("TweetData");
const container = database.container("Tweets");
const webhook = new IncomingWebhook(url);

app.get('/', async (req, res) => {
    res.sendFile(__dirname + '/index.html')
});

app.post('/api/favorite', async (req, res) => {
    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify("Marked as Favorite"));
});

app.post('/api/directmessage', async (req, res) => {
    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify("DM sent to the user"));
});

app.get('/api/data', async (req, res) => {
    const querySpec = {
        query: "SELECT  c.tweet.created_at, c.tweet.id_str, c.tweet.text, c.tweet.user.screen_name, c.sentiment.sentiment FROM c ORDER BY c._ts DESC"
    };
    const { resources: items } = await container.items
        .query(querySpec)
        .fetchAll();

    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify(items));
});

app.get('/api/piechartdata', async (req, res) => {
    const piechartquerySpec = {
        query: "SELECT COUNT(1) AS TweetCount, c.sentiment.sentiment AS TweetSentiment FROM c GROUP BY c.sentiment.sentiment"
    };
    const { resources: piechartItems } = await container.items
        .query(piechartquerySpec)
        .fetchAll();

    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify(piechartItems));
});

app.get('/api/barchartdata', async (req, res) => {
    const barchartquerySpec = {
        query: "SELECT COUNT(1) AS TweetCount, c.tweet.user.screen_name AS TweetUser FROM c GROUP BY c.tweet.user.screen_name"
    };
    const { resources: barchartItems } = await container.items
        .query(barchartquerySpec)
        .fetchAll();

    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify(barchartItems));
});

var stream = T.stream('statuses/filter', { track: '#azure', language: 'en' })
stream.on('tweet', async (tweet) => {
    let tweetText = tweet.text;
    let id = tweet.id_str;
    const documents = [tweetText];
    let results = await client.analyzeSentiment(documents);
    const result = results[0];
    if (result.error === undefined) {
        console.log("Overall sentiment:", result.sentiment);
        console.log("Scores:", result.confidenceScores);
        let newItem = { tweet: tweet, sentiment: result };
        await container.items.create(newItem);

        if (result.sentiment === "negative") {
            let username = tweet.user.screen_name;
            let tweetId = id;
            console.log('Raised an issue: ', `https://twitter.com/${username}/status/${tweetId}`);
            await webhook.send(JSON.stringify({
                "@type": "MessageCard",
                "@context": "https://schema.org/extensions",
                "summary": "Some one raised an issue.",
                "themeColor": "0078D7",
                "title": `Tweet : https://twitter.com/${username}/status/${tweetId}`,
                "sections": [
                    {
                        "activityTitle": username,
                        "activitySubtitle": tweet.created_at,
                        "text": tweet.text
                    }
                ]
            }));
        }
        io.emit('dashboardupdate', {});
    } else {
        console.error("Encountered an error:", result.error);
    }
});

const listener = server.listen(20202, function () {
    console.log('Your app is listening on port ' + listener.address().port);
});