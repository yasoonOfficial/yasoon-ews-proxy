import { Appointment, AppointmentSchema, BasePropertySet, ExchangeService, ExchangeVersion, ItemId, PropertyDefinitionBase, PropertySet, TimeZoneInfo, Uri } from "ews-javascript-api";
import { mapAppointmentToApiEvent } from "..";
import { Environment } from "../model/proxy";
import { applyCredentials } from "../proxy/helper";

export interface GetSingleCalendarEventParams {
    eventId: string;
    additionalProperties?: PropertyDefinitionBase[];
}

export class GetSingleCalendarEventRequest {

    async execute(env: Environment, params: GetSingleCalendarEventParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
        let eventId = params.eventId;

        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

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
