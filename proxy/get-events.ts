import * as moment from 'moment-timezone';

import { Environment } from "../model/proxy";
import { ExchangeService, Uri, FolderId, WellKnownFolderName, Mailbox, BasePropertySet, PropertySet, DateTime, CalendarView, AppointmentSchema, Appointment, ExchangeVersion, TimeZoneInfo } from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { mapAppointmentToApiEvent } from '../proxy/mapper';

export interface GetEventsParams {
    email: string;
    calendarId: string;
    startDate: string;
    endDate: string;
}

export class GetEventsRequest {

    async execute(env: Environment, params: GetEventsParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let userEmail = params.email;
        let calendarId = params.calendarId;
        let start = new DateTime(moment(params.startDate));
        let end = new DateTime(moment(params.endDate));
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
            if (ewsResult.Items.length === 0)
                return [];

            let itemResponse = await service.BindToItems(ewsResult.Items.map(i => i.Id), PropertySet.FirstClassProperties);

            let responseArray = [];
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

            return responseArray;
        }
        catch (e) {
            console.log(e.message, e.toString(), e.stack);
            return [];
            //res.status(500).send({ key: 'retrieveCalendarsFailed', error: e.message });
        }
    }
}