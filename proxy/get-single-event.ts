import { Appointment, AppointmentSchema, BasePropertySet, ExchangeService, ExchangeVersion, FolderId, ItemId, Mailbox, PropertyDefinitionBase, PropertySet, TimeZoneInfo, Uri, WellKnownFolderName } from "ews-javascript-api";
import { mapAppointmentToApiEvent } from "..";
import { Environment } from "../model/proxy";
import { applyCredentials } from "../proxy/helper";

export interface GetSingleCalendarEventParams {
    email: string;
    calendarId: string;
    eventId: string;
    additionalProperties?: PropertyDefinitionBase[];
}

export class GetSingleCalendarEventRequest {

    async execute(env: Environment, params: GetSingleCalendarEventParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let userEmail = params.email;
        let calendarId = params.calendarId;
        let eventId = params.eventId;
        let ewsFolder: FolderId = null;

        if (calendarId === 'main') {
            ewsFolder = new FolderId(WellKnownFolderName.Calendar, new Mailbox(userEmail));
        } else {
            ewsFolder = new FolderId();
            ewsFolder.UniqueId = calendarId;
        }

        try {
            let propSet = new PropertySet(BasePropertySet.FirstClassProperties, AppointmentSchema.StartTimeZone, AppointmentSchema.EndTimeZone);
            let item = await Appointment.Bind(service, new ItemId(eventId), propSet);

            if (item == null)
                return null;

            return mapAppointmentToApiEvent(item);

        } catch (e) {
            console.log(e.message, e.toString(), e.stack);
            return null;
        }
    }
}
