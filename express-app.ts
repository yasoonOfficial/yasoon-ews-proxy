import * as express from 'express';
import * as bodyParser from 'body-parser';
import { SearchUserRequest } from 'proxy/search-user';
import { getEnvFromHeader, tryWrapper } from 'proxy/helper';
import { GetUserRequest } from 'proxy/get-user';
import { GetUserImageRequest } from 'proxy/get-user-image';
import { GetCalendarsRequest } from 'proxy/get-calendars';
import { GetEventsRequest } from 'proxy/get-events';
import { CreateEventRequest } from 'proxy/create-event';
import { UpdateEventRequest } from 'proxy/update-event';
import { GetFreeBusyEventsRequest } from 'proxy/get-free-busy-events';
import { GetPermissionsRequest } from 'proxy/get-permissions';
import { CreateCalendarRequest } from 'proxy/create-calendar';
import { CreateWunderbarLinkRequest } from 'proxy/create-wunderbar-link';
import { GetPublicFolderMailboxRequest } from 'proxy/get-publicfolder-mailbox';
import { GetAutodiscoverDataRequest } from 'proxy/get-autodiscover-data';
import { EwsLogging } from 'ews-javascript-api';
import { EWS_AUTH_TYPE_HEADER, EWS_TOKEN_HEADER, EWS_URL_HEADER, EWS_USER_HEADER, EWS_PASSWORD_HEADER } from 'model/constants';

const customHeaders = [
    EWS_AUTH_TYPE_HEADER,
    EWS_TOKEN_HEADER,
    EWS_URL_HEADER,
    EWS_PASSWORD_HEADER,
    EWS_USER_HEADER,
];

let app = express();
app.use(bodyParser.json());
app.use(function (req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept, " + customHeaders.join(','));
    res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, PATCH");
    next();
});

EwsLogging.DebugLogEnabled = false;

app.post('/logging', (req: express.Request, res: express.Response) => {
    EwsLogging.DebugLogEnabled = req.body.enabled;
    res.status(200).send();
});

app.get('/autodiscover', tryWrapper(async (req: express.Request, res: express.Response) => {
    let userEmail = req.headers[EWS_USER_HEADER];
    let getAutodiscover = new GetAutodiscoverDataRequest();
    let result = await getAutodiscover.execute(getEnvFromHeader(req), { email: userEmail });
    res.send(result);
}));

app.get('/user/:email/publicFolderMailbox', tryWrapper(async (req: express.Request, res: express.Response) => {
    let getPublicFolder = new GetPublicFolderMailboxRequest();
    let result = await getPublicFolder.execute(getEnvFromHeader(req), req.query);
    res.send(result);
}));

app.get('/user/search', tryWrapper(async (req: express.Request, res: express.Response) => {
    let searchUser = new SearchUserRequest();
    let result = await searchUser.execute(getEnvFromHeader(req), req.query);
    res.send(result);
}));

app.get('/user/:email', tryWrapper(async (req: express.Request, res: express.Response) => {
    let getUser = new GetUserRequest();
    let result = await getUser.execute(getEnvFromHeader(req), req.query);
    res.send(result);
}));

app.get('/user/:email/photo', tryWrapper(async (req: express.Request, res: express.Response) => {
    let getUserImage = new GetUserImageRequest();
    let result = await getUserImage.execute(getEnvFromHeader(req), req.query);

    res.set('Content-Type', result.mimeType);
    res.send(result.content);
}));

app.get('/user/:email/calendars', async (req: express.Request, res: express.Response) => {
    let getUserCalendar = new GetCalendarsRequest();
    let result = await getUserCalendar.execute(getEnvFromHeader(req), req.query);
    res.send(result);
});

app.get('/user/:email/calendars/:id/events', async (req: express.Request, res: express.Response) => {
    let getUserCalendarEvents = new GetEventsRequest();
    let result = await getUserCalendarEvents.execute(getEnvFromHeader(req), {
        calendarId: req.params.id,
        email: req.params.email,
        startDate: req.query.startDate,
        endDate: req.query.endDate
    });

    res.send(result);
});

app.post('/user/:email/calendars/:id/events', tryWrapper(async (req: express.Request, res: express.Response) => {
    let createEvent = new CreateEventRequest();
    let result = await createEvent.execute(getEnvFromHeader(req), {
        calendarId: req.params.id,
        email: req.params.email
    }, req.body);

    res.send(result);
}));

app.patch('/user/:email/calendars/:id/events/:eventId', tryWrapper(async (req: express.Request, res: express.Response) => {
    let updateEvent = new UpdateEventRequest();
    await updateEvent.execute(getEnvFromHeader(req), {
        calendarId: req.params.id,
        email: req.params.email,
        eventId: req.params.eventId
    }, req.body);

    res.send(200);
}));

app.get('/user/:email/calendars/:id/free-busy', tryWrapper(async (req: express.Request, res: express.Response) => {
    if (req.params.id !== 'main')
        return res.status(400).send();

    let getUserCalendarEvents = new GetFreeBusyEventsRequest();
    let result = await getUserCalendarEvents.execute(getEnvFromHeader(req), {
        email: req.params.email,
        startDate: req.query.startDate,
        endDate: req.query.endDate
    });

    res.send(result);
}));

app.get('/user/:email/calendars/:id/effective-permissions', tryWrapper(async (req: express.Request, res: express.Response) => {
    let getPermissions = new GetPermissionsRequest();
    let result = await getPermissions.execute(getEnvFromHeader(req), {
        calendarId: req.params.id,
        email: req.params.email
    });

    res.send(result);
}));

app.post('/user/:email/calendars', tryWrapper(async (req: express.Request, res: express.Response) => {
    let createCalendar = new CreateCalendarRequest();
    let result = await createCalendar.execute(getEnvFromHeader(req), req.params, req.body);
    res.send(result);
}));

app.post('/user/:email/create-wunderbar-link', tryWrapper(async (req: express.Request, res: express.Response) => {
    let createWunderlink = new CreateWunderbarLinkRequest();
    await createWunderlink.execute(getEnvFromHeader(req), req.params, req.body);

    res.send({
        success: true
    });
}));

app.get('/', (req, res) => {
    res.status(200).send('You have been served. Nothing to see, please move on. <br/>The Job (⌐■_■)');
});

export = app;