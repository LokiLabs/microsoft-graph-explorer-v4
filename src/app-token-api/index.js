const express = require('express');
const bodyParser = require('body-parser');
const msal = require('@azure/msal-node');
const dotenv = require('dotenv');
dotenv.config();


const DEFAULT_PORT = 4000;
const app = express();
app.use(bodyParser.json());

app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept')
    next();
});

app.get('/getApplicationToken', async (req, res) => {
    try {
        const tenantUrl = 'https://login.microsoftonline.com' + req.query.tenantId;
        const tokenRequest = {
            scopes: ['https://graph.microsoft.com/.default'],
        };
        const msalConfig = {
            auth: {
                clientId: process.env.REACT_APP_CLIENT_ID,
                authority: tenantUrl,
                clientSecret: process.env.REACT_APP_CLIENT_SECRET,
            }
        };
        console.log(process.env.REACT_APP_CLIENT_ID);
        console.log(tenantUrl);
        const msalcca = new msal.ConfidentialClientApplication(msalConfig);
        const authResponse = await msalcca.acquireTokenByClientCredential(tokenRequest);
        res.status(200).send(authResponse);
    } catch (err) {
        console.log(err);
        res.status(400).send(err);
    }
});

app.listen(DEFAULT_PORT, () => {
    // eslint-disable-next-line no-console
    console.log(`listening at localhost:${DEFAULT_PORT}`);
});