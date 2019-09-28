// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides the provides server startup, authorization context creation, and the Web APIs of the add-in. 
*/

import * as fs from 'fs';
import * as https from 'https';
import * as path from 'path';
import * as express from 'express';
import * as bodyParser from 'body-parser';
import * as cors from 'cors';
import * as morgan from 'morgan';
import { AuthModule } from './auth';
import { MSGraphHelper} from './msgraph-helper';
import { UnauthorizedError } from './errors';

/* Set the environment to development if not set */
const env = process.env.NODE_ENV || 'development';

/* Instantiate AuthModule to assist with JWT parsing and verification, and token acquisition. */
const auth = new AuthModule(
    /* These values are required for our application to exhcange and get access to the resource data */
    /* client_id */ '8219744a-bdd3-4ded-8314-fa7a3be9d912',
    /* client_secret */ 'ef.LO0AtOJL5r-FCUMPX@Ye+aDs3dBR6',

    /* This information tells our server where to download the signing keys to validate the JWT that we received,
     * and where to get tokens. This is not configured for multi tenant; i.e., it is assumed that the source of the JWT and our application live
     * on the same tenant.
     */
    /* tenant */ 'common',
    /* stsDomain */ 'https://login.microsoftonline.com',
    /* discoveryURLsegment */ '.well-known/openid-configuration',
    /* tokenURLsegment */ '/oauth2/v2.0/token',

    /* Token is validated against the following values: */
    // Audience is the same as the client ID because, relative to the Office host, the add-in is the "resource".
    /* audience */ '8219744a-bdd3-4ded-8314-fa7a3be9d912',
    /* scopes */ ['access_as_user'],
    /* issuer */ 'https://login.microsoftonline.com/4073d839-2e7d-4816-ab22-428e06b5f61d/v2.0',
);

/* A promisified express handler to catch errors easily */
const handler =
    (callback: (req: express.Request, res: express.Response, next?: express.NextFunction) => Promise<any>) =>
        (req, res, next) => callback(req, res, next)
            .catch(error => {
                /* If the headers are already sent then resort to the built in error handler */
                if (res.headersSent) {
                    return next(error);
                }

                /**
                 * If running development environment we send the error details back.
                 * Else we send the right code and message.
                 */
                if (env === 'development') {
                    return res.status(error.code || 500).json({ error });
                }
                else {
                    return res.status(error.code || 500).send(error.message);
                }
            });

/* Create the express app and add the required middleware */
const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());
app.use(morgan('dev'));
app.use(express.static('public'));
/* Turn off caching when debugging */
app.use(function (req, res, next) {
    res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    res.header('Expires', '-1');
    res.header('Pragma', 'no-cache');
    next()
});

/**
 * If running on development env, then use the locally available certificates.
 */
if (env === 'development') {
    const cert = {
        key: fs.readFileSync(path.resolve('./dist/certs/server.key')),
        cert: fs.readFileSync(path.resolve('./dist/certs/server.crt'))
    };
    https.createServer(cert, app).listen(3000, () => console.log('Server running on 3000'));
}
else {
    /**
     * We don't use https as we are assuming the production environment would be on Azure.
     * Here IIS_NODE will handle https requests and pass them along to the node http module
     */
    app.listen(process.env.port || 1337, () => console.log(`Server listening on port ${process.env.port}`));
}

app.get('/index.html', handler(async (req, res) => {
    return res.sendfile('index.html');
}));

app.get('/api/values', handler(async (req, res) => {
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user Mail.Read Mail.ReadWrite' });
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Mail.Read', 'Mail.ReadWrite']);
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/messages", "");
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    const itemNames: string[] = [];
    const Items: string[] = graphData['value'];
    for (let item of Items){
        itemNames.push(item['name']);
    }
    // return res.json(itemNames);
    return res.json(Items);
}));


