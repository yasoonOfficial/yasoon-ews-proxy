import * as express from 'express';
import * as multer from 'multer';
import * as bodyParser from 'body-parser';
import * as url from 'url';
import * as request from 'request-promise';
import * as FormData from 'form-data';
import * as moment from 'moment-timezone';
import * as fs from 'fs';
import * as poxAutodiscover from 'autodiscover';
import * as xmlEscape from 'xml-escape';
import { EnumValues } from 'enum-values';
import { OfficeApiEvent, OfficeEventAttendee, EventAvailability } from './model/office';
import { DownloadRequest } from './model/downloadRequest';
import { PipeRequest, PipeUploadRequest } from './model/pipeRequest';
import { Office365GetAttachmentResponse } from './model/office365GetAttachmentResponse';
import {
    EwsLogging, ExchangeService, Folder, Item, Attendee, StringList, ItemId, SoapFaultDetails, AutodiscoverService,
    ExchangeVersion, OAuthCredentials, Uri, Attachment, BodyType, FileAttachment, SendInvitationsMode, ConflictResolutionMode,
    ItemAttachment, ServiceResult, PropertySet, EmailMessageSchema, FolderId, AttendeeCollection, MeetingResponseType, SendInvitationsOrCancellationsMode,
    WellKnownFolderName, FolderView, Mailbox, WebCredentials, SearchFilter, IsEqualTo, FolderSchema, Appointment, DateTimeKind,
    ExtendedProperty, ExtendedPropertyDefinition, BasePropertySet, MapiPropertyType, ItemTraversal, FindItemsResults,
    ItemView, AutodiscoverRedirectionUrlValidationCallback, UserSettingName, EmailMessage, AlternateId, IdFormat, AlternateIdBase,
    MessageBody, EmailAddress, DefaultExtendedPropertySet, Guid, FindFoldersResults, GetUserSettingsResponse, CalendarFolder,
    FolderPermission, StandardUser, FolderPermissionLevel, UserId, AttendeeInfo, AvailabilityData, TimeWindow, DateTime, CalendarView,
    GetUserAvailabilityResults, TimeZoneInfo, LegacyFreeBusyStatus, AvailabilityOptions, FreeBusyViewType, ExchangeCredentials, TraceFlags, ExchangeServiceBase, Exception, AppointmentSchema, Sensitivity, EnumHelper, AppointmentType
} from "ews-javascript-api";
import { ConfigurationApi } from "ews-javascript-api"; // add other imported objects based on your need
import { ntlmAuthXhrApi } from "./ews/CustomNtlmAuthXhrApi";

import { FindPeopleRequest } from './ews/FindPeopleRequest';
import { GetUserPhotoRequest } from './ews/GetUserPhotoRequest';
import { NtlmExchangeService } from './ews/NtlmAutodiscoverService';
import { AutodiscoverService as NtlmAutodiscoverService } from './ews/CustomAutodiscoverService';
import { SearchUserRequest } from 'proxy/search-user';
import { EWS_AUTH_TYPE_HEADER, EWS_TOKEN_HEADER, EWS_URL_HEADER, EWS_URL_OFFICE_365, EWS_USER_HEADER, EWS_PASSWORD_HEADER } from 'model/constants';
import { getEnvFromHeader } from 'proxy/helper';

const customHeaders = [
    EWS_AUTH_TYPE_HEADER,
    EWS_TOKEN_HEADER,
    EWS_URL_HEADER,
    EWS_PASSWORD_HEADER,
    EWS_USER_HEADER,
];

let app = express();
let upload = multer({ dest: '/tmp/' })
//let upload = multer({ dest: 'C:\\Windows\\Temp' })

app.use(bodyParser.json());
//app.use(bodyParser.urlencoded({ extended: true }));
app.use(function (req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept, " + customHeaders.join(','));
    res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, PATCH");
    next();
});

EwsLogging.DebugLogEnabled = false;

class LocalAttachment {
    content: Buffer;
    fileName: string;
    contentType: string;
}

app.post('/logging', (req: express.Request, res: express.Response) => {
    EwsLogging.DebugLogEnabled = req.body.enabled;
    res.status(200).send();
});

app.get('/region', (req: express.Request, res: express.Response) => {
    res.send({ region: process.env.region });
});

app.get('/autodiscover', tryWrapper(async (req: express.Request, res: express.Response) => {
    let userEmail = req.headers[EWS_USER_HEADER];
    let service: ExchangeService = new NtlmExchangeService();
    applyCredentials(service, req, res);

    await service.AutodiscoverUrl(userEmail, validateAutodiscoverRedirection);

    let discoverService: AutodiscoverService = new NtlmAutodiscoverService();
    discoverService.RedirectionUrlValidationCallback = validateAutodiscoverRedirection;
    applyCredentials(discoverService, req, res);

    let userSettings: GetUserSettingsResponse;

    userSettings = await discoverService.GetUserSettings(
        userEmail,
        UserSettingName.ExternalWebClientUrls
    );

    let owaUrl = null;
    let ewsUrl = service.Url.AbsoluteUri;
    if (userSettings.Settings && userSettings.Settings[UserSettingName.ExternalWebClientUrls]) {
        let urls: any[] = userSettings.Settings[UserSettingName.ExternalWebClientUrls].Urls;
        if (urls && urls.length > 0) {
            owaUrl = urls[0].Url;
        }
    }

    let authMode = 'ntlm';
    //Try if we should use ntlm or basic, fallback to ntlm
    try {
        //To-do move this to own function
        let testService = new ExchangeService();
        testService.Url = new Uri(ewsUrl);
        testService.XHRApi = new ntlmAuthXhrApi(userEmail, new Buffer(req.headers[EWS_PASSWORD_HEADER], 'base64').toString());
        testService.UseDefaultCredentials = true; //Bug... 

        var request = new FindPeopleRequest(testService, null);
        request.QueryString = userEmail;
        request.View = new ItemView(100);
        let response = await request.Execute();
    } catch (e) {
        authMode = 'basic';
    }

    let type = (ewsUrl === 'https://outlook.office365.com/EWS/Exchange.asmx') ? 'office365' : 'onpremise';
    res.send({ type: type, url: ewsUrl, owaUrl: owaUrl, authMode: authMode });
}));

app.get('/user/:email/publicFolderMailbox', tryWrapper(async (req: express.Request, res: express.Response) => {
    let userEmail = req.params.email;
    let service: AutodiscoverService = new NtlmAutodiscoverService();
    service.RedirectionUrlValidationCallback = validateAutodiscoverRedirection;

    applyCredentials(service, req, res);

    let userSettings: GetUserSettingsResponse;

    userSettings = await service.GetUserSettings(
        userEmail,
        UserSettingName.PublicFolderInformation
    );

    let anchorMailbox = userSettings.Settings[UserSettingName.PublicFolderInformation];

    //No public folder access?
    if (!anchorMailbox) {
        return res.send({ success: false });
    }

    poxAutodiscover.getPOXAutodiscoverValues(anchorMailbox, getCredentialsAsAuth(req), (err, data) => {
        if (!err && data && data.Autodiscover && data.Autodiscover.Response.length > 0) {
            let server = data.Autodiscover.Response[0].Account[0].Protocol.find(p => p.Type[0] === 'EXCH').Server[0];
            res.send({ success: true, anchorMailbox: userSettings.Settings[UserSettingName.PublicFolderInformation], publicFolderMailbox: server });
        } else {
            res.send({ success: false, err: err.toString() });
        }
    });
}));

app.get('/user/search', tryWrapper(async (req: express.Request, res: express.Response) => {
    let searchUser = new SearchUserRequest();
    let result = await searchUser.execute(getEnvFromHeader(req), req.query, null);
    res.send(result);
}));

app.get('/user/:email', tryWrapper(async (req: express.Request, res: express.Response) => {
    let service = new ExchangeService();
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    var request = new FindPeopleRequest(service, null);
    request.QueryString = req.params.email;
    request.View = new ItemView(100);
    let response = await request.Execute();

    //Special case for aliases
    if (response.People.length === 1) {
        res.send(response.People.map((p) => ({
            id: p.PersonaId.Id,
            displayName: p.DisplayName,
            mail: p.EmailAddress.EmailAddress,
            givenName: p.GivenName,
            surname: p.Surname
        }))[0]);
    } else {
        res.send(response.People.filter(p => p.PersonaType === 'Person' && p.EmailAddress.EmailAddress === req.params.email).map((p) => ({
            id: p.PersonaId.Id,
            displayName: p.DisplayName,
            mail: p.EmailAddress.EmailAddress,
            givenName: p.GivenName,
            surname: p.Surname
        }))[0]);
    }
}));

app.get('/user/:email/photo', tryWrapper(async (req: express.Request, res: express.Response) => {
    let service = new ExchangeService();
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    var request = new GetUserPhotoRequest(service, null);
    request.EmailAddress = req.params.email;
    request.Size = 360;

    let response = await request.Execute();
    res.set('Content-Type', response.PictureData.ContentType);
    res.send(new Buffer(response.PictureData.PictureData, 'base64'));
}));

app.get('/user/:email/calendars', async (req: express.Request, res: express.Response) => {
    let service = new ExchangeService(ExchangeVersion.Exchange2013);
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    let sharedCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(req.params.email));
    let folderView = new FolderView(1000);
    folderView.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.EffectiveRights);
    try {
        let ewsResult = await service.FindFolders(sharedCalendar, folderView);
        res.send(ewsResult.Folders.map(f => ({
            id: f.Id.UniqueId,
            name: f.DisplayName,
            access: getAccessArrayFromEffectiveRights(f.EffectiveRights)
        })));
    }
    catch (e) {
        res.send([]);
        console.log(e.message, e.toString(), e.stack);
        //res.status(500).send({ key: 'retrieveCalendarsFailed', error: e.message });
    }
});

app.get('/user/:email/calendars/:id/events', async (req: express.Request, res: express.Response) => {
    let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    let userEmail = req.params.email;
    let calendarId = req.params.id;
    let start = new DateTime(moment(req.query.startDate as string));
    let end = new DateTime(moment(req.query.endDate as string));
    let ewsFolder: FolderId = null;

    if (calendarId === 'main') {
        ewsFolder = new FolderId(WellKnownFolderName.Calendar, new Mailbox(userEmail));
    } else {
        ewsFolder = new FolderId();
        ewsFolder.UniqueId = calendarId;
    }

    let calendarView = new CalendarView(start, end);
    calendarView.MaxItemsReturned = 250;
    calendarView.PropertySet = new PropertySet(
        BasePropertySet.IdOnly,
        AppointmentSchema.Sensitivity,
        AppointmentSchema.Start,
        AppointmentSchema.End,
        AppointmentSchema.IsAllDayEvent,
        AppointmentSchema.LegacyFreeBusyStatus
    );

    try {
        let ewsResult = await service.FindAppointments(ewsFolder, calendarView);
        let itemResponse = await service.BindToItems(ewsResult.Items.map(i => i.Id), PropertySet.FirstClassProperties);

        let responseArray = [];
        let privateItems = [];
        for (let i = 0; i < itemResponse.Responses.length; i++) {
            let item: Appointment = <Appointment>itemResponse.Responses[i].Item;

            //Might be private.. Check in ewsResult
            if (!item) {
                //@ts-ignore
                if (ewsResult.Items[i] && ewsResult.Items[i].Sensitivity !== "Normal") {
                    responseArray.push(mapAppointmentToApiEvent(ewsResult.Items[i]));
                }
            } else {
                responseArray.push(mapAppointmentToApiEvent(item));
            }
        }

        res.send(responseArray);
    }
    catch (e) {
        res.send([]);
        //res.status(500).send({ key: 'retrieveCalendarsFailed', error: e.message });
    }
});

app.post('/user/:email/calendars/:id/events', tryWrapper(async (req: express.Request, res: express.Response) => {
    //Basic checks, appointments need at least subject, start date & end date
    let rawEvent: OfficeApiEvent = req.body;
    if (!rawEvent.subject || !rawEvent.start || !rawEvent.end)
        return res.status(400).send();

    let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    let userEmail = req.params.email;
    let calendarId = req.params.id;
    let targetFolderId: FolderId = null;
    let mode: SendInvitationsMode = SendInvitationsMode.SendToNone;

    if (rawEvent.attendees && rawEvent.attendees.length > 0) {
        mode = SendInvitationsMode.SendToAllAndSaveCopy;
    }

    //Get folder instance
    if (calendarId === 'main') {
        targetFolderId = new FolderId(WellKnownFolderName.Calendar, new Mailbox(userEmail));
    } else {
        targetFolderId = new FolderId();
        targetFolderId.UniqueId = calendarId;
    }

    let appointment = new Appointment(service);
    copyApiEventToAppointment(rawEvent, appointment);

    await appointment.Save(targetFolderId, mode);
    res.send({ id: appointment.Id.UniqueId });
}));

app.patch('/user/:email/calendars/:id/events/:eventId', tryWrapper(async (req: express.Request, res: express.Response) => {
    let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    let userEmail = req.params.email;
    let calendarId = req.params.id;
    let eventId = req.params.eventId;
    let targetFolderId: FolderId = null;
    let rawEvent: OfficeApiEvent = req.body;

    //Get folder instance
    if (calendarId === 'main') {
        targetFolderId = new FolderId(WellKnownFolderName.Calendar, new Mailbox(userEmail));
    } else {
        targetFolderId = new FolderId();
        targetFolderId.UniqueId = calendarId;
    }

    let appointment = await Appointment.Bind(service, new ItemId(eventId));
    copyApiEventToAppointment(rawEvent, appointment);

    let mode: SendInvitationsOrCancellationsMode = SendInvitationsOrCancellationsMode.SendToNone;
    if (rawEvent.attendees && Object.keys(rawEvent).length === 1) {
        mode = SendInvitationsOrCancellationsMode.SendToChangedAndSaveCopy;
    } else if (rawEvent.attendees && rawEvent.attendees.length > 0 && Object.keys(rawEvent).length > 1) {
        mode = SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy;
    }

    await appointment.Update(ConflictResolutionMode.AutoResolve, mode);
    res.send(200);
}));

app.get('/user/:email/calendars/:id/free-busy', tryWrapper(async (req: express.Request, res: express.Response) => {
    if (req.params.id !== 'main')
        return res.status(400).send();

    let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    //We don't have full read access, check if we can get free busy data
    let attendee = new AttendeeInfo(req.params.email);
    let startDate = new DateTime(moment(req.query.startDate as string));
    let endDate = new DateTime(moment(req.query.endDate as string));

    //Request as much information as possible, subject and location may be set!
    let options = new AvailabilityOptions();
    options.RequestedFreeBusyView = FreeBusyViewType.DetailedMerged;

    let availability = await service.GetUserAvailability([attendee], new TimeWindow(startDate, endDate), AvailabilityData.FreeBusy, options);
    if (availability.AttendeesAvailability.Responses[0].Result === ServiceResult.Error) {
        return res.status(500).send({ key: 'freeBusyCallFailed', error: availability.AttendeesAvailability.Responses[0].ErrorMessage })
    }

    let calendarEvents = availability.AttendeesAvailability.Responses[0].CalendarEvents;

    res.send(calendarEvents.map(c => {
        let id = 'freeBusy' + c.StartTime.ToISOString();
        let location = '';
        let subject = getFreeBusyStatusLabel(c.FreeBusyStatus);

        if (c.Details) {
            id = c.Details.StoreId || id;
            location = c.Details.Location || location;
            subject = c.Details.Subject || subject;
        }

        return <OfficeApiEvent>{
            id: id,
            calendarId: 'main',
            start: {
                dateTime: c.StartTime.ToISOString(),
                timeZone: 'UTC'
            },
            end: {
                dateTime: c.EndTime.ToISOString(),
                timeZone: 'UTC'
            },
            subject: subject,
            location: { displayName: location },
            isAllDay: (c.EndTime.Subtract(c.StartTime).TotalHours >= 24),
            showAs: getFreeBusyStatusNewName(c.FreeBusyStatus)
        };
    }));

}));

app.get('/user/:email/calendars/:id/effective-permissions', tryWrapper(async (req: express.Request, res: express.Response) => {
    if (req.params.id !== 'main')
        return res.status(400).send();

    let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    let mainCalendarId = new FolderId(WellKnownFolderName.Calendar, new Mailbox(req.params.email));
    let mainCalendar: Folder;

    //First check if we have read access to this calendar
    try {
        mainCalendar = await Folder.Bind(service, mainCalendarId, new PropertySet(BasePropertySet.IdOnly, FolderSchema.EffectiveRights));
    }
    catch (e) {
        console.log(e.message);
    }

    if (mainCalendar && mainCalendar.Id && mainCalendar.Id.UniqueId) {
        //We got at least full read access, check the rest
        if (mainCalendar.EffectiveRights) {
            let access = getAccessArrayFromEffectiveRights(mainCalendar.EffectiveRights);
            //If there is no access as of now, we only have free-busy access! Not sure how we were able to get the
            // folder ID that way, but well... Happens for timur :D
            if (access.length > 0) {
                return res.send({
                    id: mainCalendar.Id.UniqueId,
                    access: access
                });
            } //Else -> Fall back to free-busy, see below
        } else {
            //Fallback to old logic, not sure this is correct though
            return res.send({
                id: mainCalendar.Id.UniqueId,
                access: ['read']
            });
        }
    }

    //We don't have full read access, check if we can get free busy data
    let attendee = new AttendeeInfo(req.params.email);
    let availability: GetUserAvailabilityResults;

    try {
        availability = await service.GetUserAvailability([attendee], new TimeWindow(DateTime.Now, DateTime.Now.AddDays(1)), AvailabilityData.FreeBusy);
    } catch (e) {
        console.log(e.message);
        console.log(e.stack);
    }

    if (availability && availability.AttendeesAvailability && availability.AttendeesAvailability.Responses &&
        availability.AttendeesAvailability.Responses[0].Result !== ServiceResult.Error) {
        return res.send({
            access: ['freebusy']
        });
    }

    res.send({
        access: ['none']
    });
}));

app.get('/user/:email/calendars/:id/permissions', tryWrapper(async (req: express.Request, res: express.Response) => {
    if (req.params.id !== 'main')
        return res.status(400).send();

    let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    let mainCalendarId = new FolderId(WellKnownFolderName.Calendar, new Mailbox(req.params.email));
    let folderView = new FolderView(1000);
    let mainCalendar: Folder;

    //First check if we have read access to this calendar
    try {
        mainCalendar = await Folder.Bind(service, mainCalendarId, PropertySet.IdOnly);
    }
    catch (e) {
        console.log(e.message);
    }

    if (mainCalendar && mainCalendar.Id && mainCalendar.Id.UniqueId) {
        return res.send({
            id: mainCalendar.Id.UniqueId,
            access: 'read'
        });
    }

    //We don't have full read access, check if we can get free busy data
    let attendee = new AttendeeInfo(req.params.email);
    let availability: GetUserAvailabilityResults;

    try {
        availability = await service.GetUserAvailability([attendee], new TimeWindow(DateTime.Now, DateTime.Now.AddDays(1)), AvailabilityData.FreeBusy);
    } catch (e) {
        console.log(e.message);
        console.log(e.stack);
    }

    if (availability && availability.AttendeesAvailability && availability.AttendeesAvailability.Responses &&
        availability.AttendeesAvailability.Responses[0].Result !== ServiceResult.Error) {
        return res.send({
            access: 'freebusy'
        });
    }

    res.send({
        access: 'none'
    });
}));

app.post('/user/:email/calendars', tryWrapper(async (req: express.Request, res: express.Response) => {
    let service = new ExchangeService(ExchangeVersion.Exchange2013);
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    let calendarData = req.body;
    let folder = new CalendarFolder(service);
    folder.DisplayName = calendarData.name;

    let parentFolder = new FolderId(WellKnownFolderName.Calendar, new Mailbox(req.params.email));
    await folder.Save(parentFolder);

    //Yeah.. Well.. Exchange
    if (calendarData.isOwnMailbox) {
        setTimeout(async () => {
            let defaultEdit = new FolderPermission(StandardUser.Default, FolderPermissionLevel.Editor);
            folder.Permissions.Add(defaultEdit);
            await folder.Update();

            res.send({
                id: folder.Id.UniqueId,
                name: folder.DisplayName
            });
        }, 5000);
    } else {
        res.send({
            id: folder.Id.UniqueId,
            name: folder.DisplayName
        });
    }
}));

app.post('/user/:email/create-wunderbar-link', async (req: express.Request, res: express.Response) => {
    let service = new ExchangeService(ExchangeVersion.Exchange2013);
    service.Url = new Uri(req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365);
    applyCredentials(service, req, res);

    let ownUserEmail = req.params.email;
    let targetMailboxEmail = req.body.targetMailboxAddress;
    let targetMailboxFolderId = req.body.targetMailboxFolderId;

    let rootFolder = new FolderId(WellKnownFolderName.Root, new Mailbox(ownUserEmail));
    let commonViewFolderView = new FolderView(1000);
    let commonViewSearchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Common Views");
    let ewsResult: FindFoldersResults;

    try {
        ewsResult = await service.FindFolders(rootFolder, commonViewSearchFilter, commonViewFolderView);
    }
    catch (e) {
        return res.status(500).send({ key: 'findCommonViews', error: e.message });
    }

    let commonViewsFolder = ewsResult.Folders[0];

    //Constants
    let PidTagWlinkAddressBookEID = new ExtendedPropertyDefinition(0x6854, MapiPropertyType.Binary);
    let PidTagWlinkFolderType = new ExtendedPropertyDefinition(0x684F, MapiPropertyType.Binary);
    let PidTagWlinkGroupName = new ExtendedPropertyDefinition(0x6851, MapiPropertyType.String);
    let pidTagEntryId = new ExtendedPropertyDefinition(4095, MapiPropertyType.Binary);
    let PidTagNormalizedSubject = new ExtendedPropertyDefinition(0x0E1D, MapiPropertyType.String);
    let PidTagWlinkType = new ExtendedPropertyDefinition(0x6849, MapiPropertyType.Integer);
    let PidTagWlinkFlags = new ExtendedPropertyDefinition(0x684A, MapiPropertyType.Integer);
    let PidTagWlinkOrdinal = new ExtendedPropertyDefinition(0x684B, MapiPropertyType.Binary);
    let PidTagWlinkSection = new ExtendedPropertyDefinition(0x6852, MapiPropertyType.Integer);
    let PidTagWlinkGroupHeaderID = new ExtendedPropertyDefinition(0x6842, MapiPropertyType.Binary);
    let PidTagWlinkSaveStamp = new ExtendedPropertyDefinition(0x6847, MapiPropertyType.Integer);
    let PidTagWlinkStoreEntryId = new ExtendedPropertyDefinition(0x684E, MapiPropertyType.Binary);
    let PidTagWlinkGroupClsid = new ExtendedPropertyDefinition(0x6850, MapiPropertyType.Binary);
    let PidTagWlinkEntryId = new ExtendedPropertyDefinition(0x684C, MapiPropertyType.Binary);
    let PidTagWlinkRecordKey = new ExtendedPropertyDefinition(0x684D, MapiPropertyType.Binary);
    let PidTagWlinkCalendarColor = new ExtendedPropertyDefinition(0x6853, MapiPropertyType.Integer);
    let PidTagWlinkROGroupType = new ExtendedPropertyDefinition(0x6892, MapiPropertyType.Integer);
    let PidTagWlinkAddressBookStoreEID = new ExtendedPropertyDefinition(0x6891, MapiPropertyType.Binary);

    //Configure Autodiscover Service
    let autodiscoverService = new NtlmAutodiscoverService();
    autodiscoverService.RedirectionUrlValidationCallback = validateAutodiscoverRedirection;

    applyCredentials(autodiscoverService, req, res);

    let ownUserSettings: GetUserSettingsResponse;
    let targetMailBoxSettings: GetUserSettingsResponse;

    try {
        ownUserSettings = await autodiscoverService.GetUserSettings(
            ownUserEmail,
            UserSettingName.InternalRpcClientServer,
            UserSettingName.UserDN
        );

        targetMailBoxSettings = await autodiscoverService.GetUserSettings(
            targetMailboxEmail,
            UserSettingName.UserDN,
            UserSettingName.InternalRpcClientServer,
            UserSettingName.UserDisplayName
        );
    }
    catch (e) {
        return res.status(500).send({ key: 'getAutodiscoverData', error: e.message });
    }

    let ownStoreId = calculateStoreId(ownUserSettings.Settings[UserSettingName.UserDN], ownUserSettings.Settings[UserSettingName.InternalRpcClientServer]);
    let targetStoreId = calculateStoreId(targetMailBoxSettings.Settings[UserSettingName.UserDN], targetMailBoxSettings.Settings[UserSettingName.InternalRpcClientServer]);

    //let abTargetABEntryId = calculateAddressBookId(targetMailBoxSettings.Settings[UserSettingName.UserDN]);

    let sharedCalFolderId = new FolderId();

    if (targetMailboxFolderId === 'main') {
        sharedCalFolderId = new FolderId(WellKnownFolderName.Calendar, new Mailbox(targetMailboxEmail));
    } else {
        sharedCalFolderId.UniqueId = targetMailboxFolderId;
    }

    let sharedCal = await CalendarFolder.Bind(service, sharedCalFolderId, new PropertySet(BasePropertySet.FirstClassProperties, [pidTagEntryId]));
    let sharedEntryId = sharedCal.ExtendedProperties.GetOrAddExtendedProperty(pidTagEntryId).Value;

    let objWunderBarLink = new EmailMessage(service);

    if (targetMailboxFolderId === 'main') {
        objWunderBarLink.Subject = targetMailBoxSettings.Settings[UserSettingName.UserDisplayName];
    } else {
        objWunderBarLink.Subject = sharedCal.DisplayName;
    }

    objWunderBarLink.ItemClass = "IPM.Microsoft.WunderBar.Link";
    //objWunderBarLink.SetExtendedProperty(PidTagWlinkAddressBookEID, Buffer.from(abTargetABEntryId, "hex"));
    objWunderBarLink.SetExtendedProperty(PidTagWlinkAddressBookStoreEID, Buffer.from(ownStoreId, "hex"));
    objWunderBarLink.SetExtendedProperty(PidTagWlinkCalendarColor, -1);
    objWunderBarLink.SetExtendedProperty(PidTagWlinkFlags, 0);
    objWunderBarLink.SetExtendedProperty(PidTagWlinkGroupName, "Shared Calendars");
    objWunderBarLink.SetExtendedProperty(PidTagWlinkFolderType, Buffer.from("0278060000000000C000000000000046", "hex"));
    objWunderBarLink.SetExtendedProperty(PidTagWlinkGroupClsid, Buffer.from("B9F0060000000000C000000000000046", "hex"));
    objWunderBarLink.SetExtendedProperty(PidTagWlinkROGroupType, -1);
    objWunderBarLink.SetExtendedProperty(PidTagWlinkSection, 3);
    objWunderBarLink.SetExtendedProperty(PidTagWlinkType, 2);
    objWunderBarLink.SetExtendedProperty(PidTagWlinkEntryId, sharedEntryId);
    objWunderBarLink.SetExtendedProperty(PidTagWlinkStoreEntryId, Buffer.from(targetStoreId, "hex"));
    objWunderBarLink.IsAssociated = true;

    try {
        await objWunderBarLink.Save(commonViewsFolder.Id);
    }
    catch (e) {
        return res.status(500).send({ key: 'createWunderbarLink', error: e.message });
    }

    res.send({
        success: true
    });
});

app.post('/downloadAttachment', async (req: express.Request, res: express.Response) => {
    let body: DownloadRequest = req.body;
    try {
        let file: LocalAttachment;
        if (body.mode === 'ews') {
            file = await handleEwsDownloadRequest(body, req, res);
        }
        else if (body.mode === 'rest') {
            file = await handleRestDownloadRequest(body, req, res);
        }

        let formData = new FormData();
        formData.append('file', file.content, { filename: file.fileName, contentType: file.contentType });
        res.setHeader('Content-Type', formData.getHeaders()['content-type']);
        res.setHeader('Content-Length', formData['getLengthSync']());
        formData.pipe(res);
    }
    catch (e) {
        console.error(e);
    }
});

app.post('/pipeAttachment', async (req: express.Request, res: express.Response) => {
    let body: PipeRequest = req.body;
    try {
        let file: LocalAttachment;
        if (body.mode === 'ews') {
            file = await handleEwsDownloadRequest(body, req, res);
        }
        else if (body.mode === 'rest') {
            file = await handleRestDownloadRequest(body, req, res);
        }

        let formData = {
            file: {
                value: file.content,
                options: {
                    filename: file.fileName,
                    contentType: file.contentType
                }
            }
        };

        request.post(`${body.baseUrl}/rest/api/2/issue/${body.issueId}/attachments`, {
            formData: formData,
            headers: {
                'X-Atlassian-Token': 'no-check',
                'Authorization': body.authHeader
            }
        }).pipe(res);
    }
    catch (e) {
        console.error(e);
    }
});

app.post('/pipeAttachmentUpload', upload.array('file'), async (req: express.Request, res: express.Response) => {
    let files: Express.Multer.File[] = <any>req.files;

    try {
        let formData = {
            file: files.map(file => ({
                value: fs.createReadStream(file.path),
                options: {
                    filename: file.originalname,
                    contentType: file.mimetype
                }
            }))
        };

        request.post(`${req.headers['x-jira-url']}/rest/api/2/issue/${req.headers['x-jira-issue-id']}/attachments`, {
            formData: formData,
            headers: {
                'X-Atlassian-Token': 'no-check',
                'Authorization': req.headers['x-jira-token']
            }
        }).pipe(res);
    }
    catch (e) {
        console.error(e);
    }
    finally {
        //Cleanup
        files.forEach(file => {
            fs.unlink(file.path);
        });
    }
});

app.get('/', (req, res) => {
    res.status(200).send('You have been served. Nothing to see, please move on. <br/>The Job (⌐■_■)');
});

function getAccessArrayFromEffectiveRights(effectiveRights: any) {
    if (effectiveRights && Number.isInteger(effectiveRights)) {
        //Todo!!!
        throw new Exception('not supported yet, todo!');
        //Bug in ews-javascript-api, see https://github.com/gautamsi/ews-javascript-api/pull/214
    } else if (effectiveRights && effectiveRights['__type'] === 'EffectiveRights') {
        let rights: {
            CreateAssociated: "true" | "false",
            CreateContents: "true" | "false",
            CreateHierarchy: "true" | "false",
            Delete: "true" | "false",
            Modify: "true" | "false",
            Read: "true" | "false",
            ViewPrivateItems: "true" | "false"
        } = <any>effectiveRights;

        let access = [];
        if (rights.CreateContents === 'true')
            access.push('create');
        if (rights.CreateHierarchy === 'true')
            access.push('createFolder');
        if (rights.Delete === 'true')
            access.push('delete');
        if (rights.Modify === 'true')
            access.push('edit');
        if (rights.Read === 'true')
            access.push('read');

        return access;
    } else {
        return [];
    }
}

function tryWrapper(func: (req: express.Request, res: express.Response) => Promise<any>): (req: express.Request, res: express.Response) => void {
    return (async (req: express.Request, res: express.Response) => {
        try {
            await func(req, res);
        }
        catch (e) {
            if (e instanceof SoapFaultDetails) {
                res.status(e.HttpStatusCode).send(e.Message);
            } else {
                res.status(500).send();
            }

            if (e && e.message) {
                console.log(e.message);
            }
            if (e && e.stack) {
                console.log(e.stack);
            }
            if (e && e.toString) {
                console.log(e.toString());
            }
        }
    });
}

function copyApiEventToAppointment(rawEvent: OfficeApiEvent, appointment: Appointment) {
    //Let the mapping begin!
    if (rawEvent.attendees) {
        rawEvent.attendees.forEach(a => {
            if (!a.type || a.type === 'Required') {
                appointment.RequiredAttendees.Add(a.emailAddress.address);
            } else if (a.type === 'Optional') {
                appointment.OptionalAttendees.Add(a.emailAddress.address);
            } else if (a.type === 'Resource') {
                appointment.Resources.Add(a.emailAddress.address);
            }
        });
    }

    if (rawEvent.body) {
        let bodyType = BodyType.Text;
        if (rawEvent.body.contentType.toLowerCase() === 'html')
            bodyType = BodyType.HTML;

        //#shitty api
        if (bodyType === BodyType.HTML)
            rawEvent.body.content = xmlEscape(rawEvent.body.content);

        appointment.Body = new MessageBody(bodyType, rawEvent.body.content);
    }

    if (rawEvent.showAs) {
        //@ts-ignore
        appointment.LegacyFreeBusyStatus = EnumValues.getNameFromValue(LegacyFreeBusyStatus, getLegacyFreeBusyStatusFromName(rawEvent.showAs));
    }

    if (rawEvent.categories) {
        appointment.Categories = new StringList(rawEvent.categories || []);
    }

    if (rawEvent.location) {
        appointment.Location = rawEvent.location ? rawEvent.location.displayName : '';
    }

    if (rawEvent.subject) {
        appointment.Subject = rawEvent.subject;
    }

    if (rawEvent.isAllDay !== undefined) {
        appointment.IsAllDayEvent = rawEvent.isAllDay;
    }

    if (rawEvent.isReminderOn === true) {
        appointment.IsReminderSet = true;
        appointment.ReminderMinutesBeforeStart = rawEvent.reminderMinutesBeforeStart;
    } else if (rawEvent.isReminderOn === false) {
        appointment.IsReminderSet = false;
    }

    if (rawEvent.start) {
        let mDate = moment.tz(rawEvent.start.dateTime as string, rawEvent.start.timeZone);
        appointment.Start = new DateTime(mDate);
        appointment.Start.kind = DateTimeKind.Unspecified;
        //appointment.StartTimeZone = TimeZoneInfo.FindSystemTimeZoneById(rawEvent.start.timeZone);
    }

    if (rawEvent.end) {
        let mDate = moment.tz(rawEvent.end.dateTime as string, rawEvent.end.timeZone);
        appointment.End = new DateTime(mDate);
        appointment.End.kind = DateTimeKind.Unspecified;
        //appointment.EndTimeZone = TimeZoneInfo.FindSystemTimeZoneById(rawEvent.end.timeZone);
    }
}

function mapAppointmentToApiEvent(item: Appointment): OfficeApiEvent {
    if (!item)
        return null;

    let attendees = mapAttendees(item.RequiredAttendees, "Required");
    let withOptional = attendees.concat(mapAttendees(item.OptionalAttendees, "Optional"));
    let all = withOptional.concat(mapAttendees(item.Resources, "Resource"));

    let result: OfficeApiEvent = null;

    //@ts-ignore
    if (item.Sensitivity !== "Normal") {
        result = {
            id: item.Id.UniqueId,
            start: {
                dateTime: item.Start.ToISOString(),
                timeZone: 'UTC'
            },
            end: {
                dateTime: item.End.ToISOString(),
                timeZone: 'UTC'
            },
            type: 'singleInstance',
            isAllDay: item.IsAllDayEvent,
            sensitivity: <any>item.Sensitivity,
            subject: <any>item.Sensitivity,
            //@ts-ignore
            showAs: getFreeBusyStatusNewName(LegacyFreeBusyStatus[item.LegacyFreeBusyStatus])
        };
    } else {
        result = {
            id: item.Id.UniqueId,
            calendarId: (item.ParentFolderId ? item.ParentFolderId.UniqueId : ''),
            subject: item.Subject,
            start: {
                dateTime: item.Start.ToISOString(),
                timeZone: 'UTC'
            },
            end: {
                dateTime: item.End.ToISOString(),
                timeZone: 'UTC'
            },
            location: { displayName: item.Location },
            isAllDay: item.IsAllDayEvent,
            //@ts-ignore
            showAs: getFreeBusyStatusNewName(LegacyFreeBusyStatus[item.LegacyFreeBusyStatus]),
            categories: (item.Categories ? item.Categories.GetEnumerator() : []),
            organizer: (item.Organizer ? ({
                emailAddress: {
                    name: item.Organizer.Name,
                    address: item.Organizer.Address
                }
            }) : null),
            type: getAppointmentType(<any>item.AppointmentType),
            isReminderOn: item.IsReminderSet,
            reminderMinutesBeforeStart: (item.IsReminderSet) ? item.ReminderMinutesBeforeStart : undefined,
            attendees: all,
            sensitivity: <any>item.Sensitivity,
            body: (item.Body ? ({
                contentType: EnumValues.getNameFromValue(BodyType, item.Body.BodyType),
                content: item.Body.Text
            }) : null)
        };
    }

    return result;
}

function validateAutodiscoverRedirection(redirectionUrl: string) {
    //Todo
    //return redirectionUrl === 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc';
    return true;
}

function mapAttendees(attendees: AttendeeCollection, type: string) {
    return attendees.GetEnumerator().map(a => ({
        type: type,
        status: {
            response: getResponseStatusName(a.ResponseType),
            time: a.LastResponseTime ? a.LastResponseTime.ToISOString() : null
        },
        emailAddress: {
            name: a.Name,
            address: a.Address
        }
    }));
}

function applyCredentials(service: ExchangeServiceBase, request: express.Request, response: express.Response) {
    if (request.headers[EWS_TOKEN_HEADER]) {
        service.Credentials = new OAuthCredentials(request.headers[EWS_TOKEN_HEADER]);
    } else if (request.headers[EWS_USER_HEADER] && request.headers[EWS_PASSWORD_HEADER] && request.headers[EWS_AUTH_TYPE_HEADER] === 'ntlm') {
        let userEmail = request.headers[EWS_USER_HEADER];
        let password = new Buffer(request.headers[EWS_PASSWORD_HEADER], 'base64').toString();
        service.XHRApi = new ntlmAuthXhrApi(userEmail, password);
        service.UseDefaultCredentials = true; //Bug... 
    } else if (request.headers[EWS_USER_HEADER] && request.headers[EWS_PASSWORD_HEADER]) {
        let userEmail = request.headers[EWS_USER_HEADER];
        let password = new Buffer(request.headers[EWS_PASSWORD_HEADER], 'base64').toString();
        service.Credentials = new WebCredentials(userEmail, password);
    }
    else {
        response.status(500).send();
        throw new Error('No Auth!');
    }
}

function getCredentialsAsAuth(request: express.Request) {
    if (request.headers[EWS_TOKEN_HEADER]) {
        return 'Bearer ' + request.headers[EWS_TOKEN_HEADER];
    }
    else if (request.headers[EWS_USER_HEADER] && request.headers[EWS_PASSWORD_HEADER]) {
        let password = new Buffer(request.headers[EWS_PASSWORD_HEADER], 'base64').toString();
        return 'Basic ' + Buffer.from(request.headers[EWS_USER_HEADER] + ':' + password).toString('base64');
    }
}

function getAppointmentType(type: string) {
    switch (type) {
        case "Single":
            return "singleInstance";
        case "Occurrence":
            return "occurrence";
        case "Exception":
            return "exception";
        case "RecurringMaster":
            return "seriesMaster";
    }
}

function getResponseStatusName(type: MeetingResponseType) {
    switch (type) {
        case MeetingResponseType.Accept:
            return "Accepted";
        case MeetingResponseType.Decline:
            return "Declined";
        case MeetingResponseType.NoResponseReceived:
            return "NotResponded";
        case MeetingResponseType.Organizer:
            return "Organizer";
        case MeetingResponseType.Tentative:
            return "TentativelyAccepted";
        case MeetingResponseType.Unknown:
            return "None";
    }
}

function getFreeBusyStatusLabel(status: LegacyFreeBusyStatus): string {
    switch (status) {
        case LegacyFreeBusyStatus.Busy:
            return 'Busy';
        case LegacyFreeBusyStatus.Free:
            return 'Free';
        case LegacyFreeBusyStatus.NoData:
            return 'Unknown';
        case LegacyFreeBusyStatus.OOF:
            return 'Out of Office';
        case LegacyFreeBusyStatus.Tentative:
            return 'Tentative';
        case LegacyFreeBusyStatus.WorkingElsewhere:
            return 'Working Elsewhere';
    }
}


function getLegacyFreeBusyStatusFromName(name: string): LegacyFreeBusyStatus {
    switch (name) {
        case 'busy':
            return LegacyFreeBusyStatus.Busy;
        case 'free':
            return LegacyFreeBusyStatus.Free;
        case 'unknown':
            return LegacyFreeBusyStatus.NoData;
        case 'oof':
            return LegacyFreeBusyStatus.OOF;
        case 'tentative':
            return LegacyFreeBusyStatus.Tentative;
        case 'workingElsewhere':
            return LegacyFreeBusyStatus.WorkingElsewhere;
    }
}

function getFreeBusyStatusNewName(status: LegacyFreeBusyStatus): EventAvailability {
    if (status === null)
        return null;

    switch (status) {
        case LegacyFreeBusyStatus.Busy:
            return <EventAvailability>'busy';
        case LegacyFreeBusyStatus.Free:
            return <EventAvailability>'free';
        case LegacyFreeBusyStatus.NoData:
            return <EventAvailability>'unknown';
        case LegacyFreeBusyStatus.OOF:
            return <EventAvailability>'oof';
        case LegacyFreeBusyStatus.Tentative:
            return <EventAvailability>'tentative';
        case LegacyFreeBusyStatus.WorkingElsewhere:
            return <EventAvailability>'workingElsewhere';
    }
}

function calculateStoreId(userDn: string, serverName: string) {
    let userDnHex = ''
    for (let i = 0; i < userDn.length; i++)
        userDnHex += (userDn.charCodeAt(i) >>> 0).toString(16).toUpperCase();

    let serverNameHex = '';
    for (let i = 0; i < serverName.length; i++)
        serverNameHex += (serverName.charCodeAt(i) >>> 0).toString(16).toUpperCase();

    let flags = "00000000";
    let ProviderUID = "38A1BB1005E5101AA1BB08002B2A56C2";
    let versionFlag = "0000";
    let DLLFileName = "454D534D44422E444C4C00000000";
    let WrappedFlags = "00000000";
    let WrappedProviderUID = "1B55FA20AA6611CD9BC800AA002FC45A";
    let WrappedType = "0C000000";
    let StoredIdStringHex = flags + ProviderUID + versionFlag + DLLFileName + WrappedFlags + WrappedProviderUID + WrappedType + serverNameHex + "00" + userDnHex + "00";
    return StoredIdStringHex;
    /*
    let sender = "confluence@yasoon.com";
    let something = '';
    for (let i = 0; i < sender.length; i++)
        something = (sender.charCodeAt(i) >>> 0).toString(16).toUpperCase() + '00';
 
    return StoredIdStringHex + 'E94632F4480000000200000010000000' + something + '00000000';*/
}

function calculateAddressBookId(userDn: string) {
    let userDnHex = ''
    for (let i = 0; i < userDn.length; i++)
        userDnHex += (userDn.charCodeAt(i) >>> 0).toString(16).toUpperCase();

    let Provider = "00000000DCA740C8C042101AB4B908002B2FE1820100000000000000";
    let userdnStringHex = Provider + userDnHex + "00";
    return userdnStringHex;
}

function guidToBytes(guid) {
    var bytes = [];
    guid.split('-').map((number, index) => {
        var bytesInChar = index < 3 ? number.match(/.{1,2}/g).reverse() : number.match(/.{1,2}/g);
        bytesInChar.map((byte) => { bytes.push(parseInt(byte, 16)); })
    });
    return bytes;
}

async function handleEwsDownloadRequest(body: DownloadRequest, req: express.Request, res: express.Response): Promise<LocalAttachment> {
    var exch = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
    exch.Credentials = new OAuthCredentials(body.ewsToken);
    exch.Url = new Uri(body.ewsUrl || EWS_URL_OFFICE_365);

    try {
        let response = await exch.GetAttachments([body.attachmentId], BodyType.HTML, [EmailMessageSchema.MimeContent]);

        if (response.Responses.length === 0 || response.Responses[0].Result != ServiceResult.Success) {
            if (response.Responses.length > 0) {
                let result = response.Responses[0];

                let ownUserSettings: GetUserSettingsResponse;
                let targetMailBoxSettings: GetUserSettingsResponse;
                console.log("Error downloading attachment", result.ErrorCode, result.ErrorMessage);
            }

            res.status(500);
        } else if (response.Responses[0].Attachment instanceof FileAttachment) {
            let file = response.Responses[0].Attachment as FileAttachment;
            return {
                content: new Buffer(file.Base64Content, "base64"),
                fileName: body.fileName,
                contentType: file.ContentType
            };
        } else if (response.Responses[0].Attachment instanceof ItemAttachment) {
            let mail = response.Responses[0].Attachment as ItemAttachment;
            return {
                content: new Buffer(mail.Item.MimeContent.Content, "base64"),
                fileName: body.fileName,
                contentType: 'message/rfc822'
            };
        }
    }
    catch (e) {
        console.log("Error while getting attachments", e, e.stack);
    }
}

async function handleRestDownloadRequest(body: DownloadRequest, req: express.Request, res: express.Response): Promise<LocalAttachment> {
    try {
        let baseUrl = body.restUrl;

        if (!baseUrl.endsWith("/"))
            baseUrl += "/";

        let url = `${baseUrl}v2.0/me/messages/${body.messageId}/attachments/${body.attachmentId}`;
        let response: Office365GetAttachmentResponse = await request.get(url, {
            json: true,
            auth: {
                'bearer': body.restToken
            }
        });

        if (response['@odata.type'] === '#Microsoft.OutlookServices.FileAttachment') {
            return {
                content: new Buffer(response.ContentBytes, "base64"),
                fileName: body.fileName,
                contentType: response.ContentType
            };
        } else if (response['@odata.type'] === '#Microsoft.OutlookServices.ItemAttachment') {
            return {
                content: new Buffer(response.ContentBytes, "base64"),
                fileName: body.fileName,
                contentType: 'message/rfc822'
            };
        }
    }
    catch (e) {
        console.log("Error while getting attachments", e, e.stack);
    }
}

export = app;