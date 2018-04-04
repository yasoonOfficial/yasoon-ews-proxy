import { Environment } from "../model/proxy";
import { ExchangeService, Uri, FolderId, WellKnownFolderName, Mailbox, ExchangeVersion, TimeZoneInfo, PropertyDefinitionBase, ItemId, Appointment } from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { mapAppointmentToApiEvent } from "..";

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
            let item = await Appointment.Bind(service, new ItemId(eventId));

            if (item == null)
                return null;

            return mapAppointmentToApiEvent(item);

        } catch (e) {
            console.log(e.message, e.toString(), e.stack);
            return null;
        }
    }
}
