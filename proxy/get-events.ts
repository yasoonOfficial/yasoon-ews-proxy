import * as moment from 'moment-timezone';

import { Environment } from "../model/proxy";
import { ExchangeService, Uri, FolderId, WellKnownFolderName, Mailbox, BasePropertySet, PropertySet, DateTime, CalendarView, AppointmentSchema, Appointment, ExchangeVersion, TimeZoneInfo, PropertyDefinitionBase } from "ews-javascript-api";
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
        additionalProperties.push(AppointmentSchema.Sensitivity);
        additionalProperties.push(AppointmentSchema.Start);
        additionalProperties.push(AppointmentSchema.End);
        additionalProperties.push(AppointmentSchema.IsAllDayEvent);
        additionalProperties.push(AppointmentSchema.LegacyFreeBusyStatus);

        calendarView.PropertySet = new PropertySet(BasePropertySet.IdOnly, additionalProperties);

        try {
            let ewsResult = await service.FindAppointments(ewsFolder, calendarView);
            if (ewsResult.Items.length === 0)
                return [];

            let props = new PropertySet(BasePropertySet.FirstClassProperties, AppointmentSchema.StartTimeZone);
            let itemResponse = await service.BindToItems(ewsResult.Items.map(i => i.Id), props);

            let responseArray: OfficeApiEvent[] = [];
            for (let i = 0; i < itemResponse.Responses.length; i++) {
                let item: Appointment = <Appointment>itemResponse.Responses[i].Item;

                //Might be private.. Check in ewsResult
                if (!item) {
                    //@ts-ignore
                    if (ewsResult.Items[i] && ewsResult.Items[i].Sensitivity !== "Normal") {
                        responseArray.push(mapAppointmentToApiEvent(ewsResult.Items[i], params.additionalProperties));
                    }
                } else {
                    responseArray.push(mapAppointmentToApiEvent(item, params.additionalProperties));
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