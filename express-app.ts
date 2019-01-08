import * as bodyParser from 'body-parser';
import { EwsLogging } from 'ews-javascript-api';
import * as express from 'express';
import { Monkey } from './extensions/Monkey';
import { EWS_AUTH_TYPE_HEADER, EWS_PASSWORD_HEADER, EWS_TOKEN_HEADER, EWS_URL_HEADER, EWS_USER_HEADER, PROXY_SECRET_HEADER } from './model/constants';
import { CreateCalendarRequest } from './proxy/create-calendar';
import { CreateEventRequest } from './proxy/create-event';
import { CreateGroupRequest } from './proxy/create-group';
import { CreateWunderbarLinkRequest } from './proxy/create-wunderbar-link';
import { DeleteCalendarRequest } from './proxy/delete-calendar';
import { DeleteEventRequest } from './proxy/delete-event';
import { FindGroupRequest } from './proxy/find-group';
import { GetAutodiscoverDataRequest } from './proxy/get-autodiscover-data';
import { GetCalendarsRequest } from './proxy/get-calendars';
import { GetCategoriesRequest } from './proxy/get-categories';
import { GetEventsRequest } from './proxy/get-events';
import { GetFreeBusyEventsRequest } from './proxy/get-free-busy-events';
import { GetOwnUserRequest } from './proxy/get-own-user';
import { GetPermissionsRequest } from './proxy/get-permissions';
import { GetPublicFolderMailboxRequest } from './proxy/get-publicfolder-mailbox';
import { GetSingleCalendarEventRequest } from './proxy/get-single-event';
import { GetUserRequest } from './proxy/get-user';
import { GetUserImageRequest } from './proxy/get-user-image';
import { decodeUrlId, requestWrapper } from './proxy/helper';
import { SearchUserRequest } from './proxy/search-user';
import { UpdateEventRequest } from './proxy/update-event';
//@ts-ignore
import * as version from './version.json';
import { Environment } from './model/proxy';

const customHeaders = [
    EWS_AUTH_TYPE_HEADER,
    EWS_TOKEN_HEADER,
    EWS_URL_HEADER,
    EWS_PASSWORD_HEADER,
    EWS_USER_HEADER,
    PROXY_SECRET_HEADER
];

let app = express();
let router = express.Router();
app.use(bodyParser.json());
app.use(function (req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept, " + customHeaders.join(','));
    res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, PATCH");
    next();
});

let secret = '';
app['configureApp'] = (s, l) => {
    secret = s;
    EwsLogging.DebugLogEnabled = !!l;
};

EwsLogging.DebugLogEnabled = false;

//Fix for https://github.com/gautamsi/ews-javascript-api/pull/219
new Monkey().patch();

router.post('/logging', (req: express.Request, res: express.Response) => {
    EwsLogging.DebugLogEnabled = req.body.enabled;
    res.status(200).send();
});

router.get('/autodiscover/:email', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getAutodiscover = new GetAutodiscoverDataRequest();
    let result = await getAutodiscover.execute(env, req.params);
    res.send(result);
}));

router.get('/user/:email/publicFolderMailbox', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getPublicFolder = new GetPublicFolderMailboxRequest();
    let result = await getPublicFolder.execute(env, req.params);
    res.send(result);
}));

router.get('/user/search', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let searchUser = new SearchUserRequest();
    let result = await searchUser.execute(env, req.query);
    res.send(result);
}));

router.get('/user/me', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getUser = new GetOwnUserRequest();
    let result = await getUser.execute(env);
    res.send(result);
}));

router.get('/user/:email', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getUser = new GetUserRequest();
    let result = await getUser.execute(env, req.params);
    res.send(result);
}));

router.get('/user/:email/photo', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getUserImage = new GetUserImageRequest();
    let result = await getUserImage.execute(env, req.params);

    res.set('Content-Type', result.mimeType);
    res.send(result.content);
}));

router.get('/user/:email/calendars', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getUserCalendar = new GetCalendarsRequest();
    let result = await getUserCalendar.execute(env, req.params);
    res.send(result);
}));

router.get('/user/:email/categories', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getCategories = new GetCategoriesRequest();
    let result = await getCategories.execute(env, req.params);
    res.send(result);
}));

router.get('/user/:email/calendars/:id/events', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getUserCalendarEvents = new GetEventsRequest();
    let result = await getUserCalendarEvents.execute(env, {
        calendarId: decodeUrlId(req.params.id),
        email: req.params.email,
        startDate: req.query.startDate,
        endDate: req.query.endDate
    });

    res.send(result);
}));

router.get('/user/:email/calendars/:id/events/:eventId', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getSingleCalendarEvent = new GetSingleCalendarEventRequest();
    let result = await getSingleCalendarEvent.execute(env, {
        eventId: decodeUrlId(req.params.eventId),
    });

    res.send(result);
}));

router.post('/user/:email/calendars/:id/events', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let createEvent = new CreateEventRequest();
    let result = await createEvent.execute(env, {
        calendarId: decodeUrlId(req.params.id),
        email: req.params.email,
        skipSendingInviteToGroup: req.query.skipSendingInviteToGroup === "true"
    }, req.body);

    res.send(result);
}));

router.patch('/user/:email/calendars/:id/events/:eventId', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let updateEvent = new UpdateEventRequest();
    await updateEvent.execute(env, {
        calendarId: decodeUrlId(req.params.id),
        email: req.params.email,
        eventId: decodeUrlId(req.params.eventId),
        entireSeries: req.query.entireSeries === 'true'
    }, req.body);

    res.status(200).send({});
}));

router.post('/user/:email/calendars/:id/events/:eventId/delete', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let deleteRequest = new DeleteEventRequest();
    await deleteRequest.execute(env, {
        eventId: decodeUrlId(req.params.eventId),
        sendCancellations: req.body.sendCancellations,
        entireSeries: req.body.entireSeries,
        cancellationMessage: req.body.cancellationMessage,
        type: req.body.type,
        doHardDelete: req.body.doHardDelete ? req.body.doHardDelete : false
    });
    res.status(200).send({});
}));

router.get('/user/:email/calendars/:id/free-busy', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    if (req.params.id !== 'main')
        return res.status(400).send();

    let getUserCalendarEvents = new GetFreeBusyEventsRequest();
    let result = await getUserCalendarEvents.execute(env, {
        email: req.params.email,
        startDate: req.query.startDate,
        endDate: req.query.endDate
    });

    res.send(result);
}));

router.get('/user/:email/calendars/:id/effective-permissions', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let getPermissions = new GetPermissionsRequest();
    let result = await getPermissions.execute(env, {
        calendarId: decodeUrlId(req.params.id),
        email: req.params.email
    });

    res.send(result);
}));

router.post('/user/:email/calendars', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let createCalendar = new CreateCalendarRequest();
    let result = await createCalendar.execute(env, req.params, req.body);
    res.send(result);
}));

router.delete('/user/:email/calendars/:id/delete', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let deleteRequest = new DeleteCalendarRequest();
    await deleteRequest.execute(env, {
        calendarId: decodeUrlId(req.params.id),
        email: req.params.email
    });
    res.status(204).send({});
}));

router.post('/user/:email/create-wunderbar-link', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let createWunderlink = new CreateWunderbarLinkRequest();
    await createWunderlink.execute(env, req.params, req.body);

    res.send({
        success: true
    });
}));

router.get('/groups', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let findGroups = new FindGroupRequest();
    let result = await findGroups.execute(env, req.query);
    res.send(result);
}));

router.post('/groups', requestWrapper(secret, async (req: express.Request, res: express.Response, env: Environment) => {
    let createGroup = new CreateGroupRequest();
    let result = await createGroup.execute(env, req.body);
    res.send(result);
}));

router.get('/', (req, res) => {
    res.status(200).send('You have been served. Nothing to see, please move on. <br/>The Job (⌐■_■)');
});

router.get('/version', (req, res) => {
    res.status(200).send(version);
});

app.use('/v2', router);
export = app;