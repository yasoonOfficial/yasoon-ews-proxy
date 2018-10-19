import { Appointment, AppointmentSchema, BasePropertySet, ExchangeService, ExchangeVersion, ItemId, PropertyDefinitionBase, PropertySet, TimeZoneInfo, Uri } from "ews-javascript-api";
import { Environment } from "../model/proxy";
import { applyCredentials } from "../proxy/helper";
import { mapAppointmentToApiEvent } from "../proxy/mapper";

export interface GetSingleCalendarEventParams {
    eventId: string;
    additionalProperties?: PropertyDefinitionBase[];
}

export class GetSingleCalendarEventRequest {

    async execute(env: Environment, params: GetSingleCalendarEventParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2010, TimeZoneInfo.Utc);
        let eventId = params.eventId;

        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let propSet = new PropertySet(BasePropertySet.FirstClassProperties, AppointmentSchema.StartTimeZone, AppointmentSchema.EndTimeZone);
        let item = await Appointment.Bind(service, new ItemId(eventId), propSet);

        if (item == null)
            return null;

        return await mapAppointmentToApiEvent(item);
    }
}
