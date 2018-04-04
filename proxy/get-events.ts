import * as moment from 'moment-timezone';

import { Environment } from "../model/proxy";
import { ExchangeService, Uri, FolderId, WellKnownFolderName, Mailbox, BasePropertySet, PropertySet, DateTime, CalendarView, AppointmentSchema, ExchangeVersion, TimeZoneInfo, PropertyDefinitionBase } from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { mapAppointmentToApiEvent } from '../proxy/mapper';
import { OfficeApiEvent } from '../model/office';

export interface GetEventsParams {
    email: string;
    calendarId: string;
    startDate: string;
    endDate: string;
    additionalProperties?: PropertyDefinitionBase[];
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

        let additionalProperties = params.additionalProperties || [];
        additionalProperties.push(AppointmentSchema.Subject);
        additionalProperties.push(AppointmentSchema.Sensitivity);
        additionalProperties.push(AppointmentSchema.Start);
        additionalProperties.push(AppointmentSchema.StartTimeZone);
        additionalProperties.push(AppointmentSchema.End);
        additionalProperties.push(AppointmentSchema.EndTimeZone);
        additionalProperties.push(AppointmentSchema.IsAllDayEvent);
        additionalProperties.push(AppointmentSchema.LegacyFreeBusyStatus);
        additionalProperties.push(AppointmentSchema.Location);
        additionalProperties.push(AppointmentSchema.Categories);
        additionalProperties.push(AppointmentSchema.ParentFolderId);
        additionalProperties.push(AppointmentSchema.AppointmentType);

        calendarView.PropertySet = new PropertySet(BasePropertySet.IdOnly, additionalProperties);

        try {
            let ewsResult = await service.FindAppointments(ewsFolder, calendarView);
            if (ewsResult.Items.length === 0)
                return [];

            let responseArray: OfficeApiEvent[] = [];
            for (let i = 0; i < ewsResult.Items.length; i++) {
                let item = ewsResult.Items[i];

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