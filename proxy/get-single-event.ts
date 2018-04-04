import { Environment } from "../model/proxy";
import { ExchangeService, Uri, FolderId, WellKnownFolderName, Mailbox, AppointmentSchema, ExchangeVersion, TimeZoneInfo, PropertyDefinitionBase, PropertySet, BasePropertySet, ItemId, Appointment } from "ews-javascript-api";
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
            let itemProps = new PropertySet(BasePropertySet.FirstClassProperties,
                AppointmentSchema.StartTimeZone, AppointmentSchema.EndTimeZone);

            let itemId = [new ItemId(eventId)];

            let itemResponse = service.BindToItems(itemId, itemProps);
            let item: Appointment = <Appointment>itemResponse[0].Responses[0].Item;

            if (item == null)
                return null;

            let additionalProperties = [];
            additionalProperties.push(AppointmentSchema.Sensitivity);
            additionalProperties.push(AppointmentSchema.Start);
            additionalProperties.push(AppointmentSchema.End);
            additionalProperties.push(AppointmentSchema.IsAllDayEvent);
            additionalProperties.push(AppointmentSchema.LegacyFreeBusyStatus);

            return mapAppointmentToApiEvent(item, additionalProperties);

        } catch (e) {
            console.log(e.message, e.toString(), e.stack);
            return null;
        }
    }
}
