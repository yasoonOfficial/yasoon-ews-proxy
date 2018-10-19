import { Environment } from "../model/proxy";
import {
    ExchangeService, Uri, Appointment,
    ExchangeVersion, TimeZoneInfo, ItemId,
    SendInvitationsOrCancellationsMode, ConflictResolutionMode
} from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { copyApiEventToAppointment } from '../proxy/mapper';
import { OfficeApiEvent } from '../model/office';

export interface UpdateUserCalendarEventParams {
    email: string;
    calendarId: string;
    eventId: string;
    entireSeries: boolean;
}

export class UpdateEventRequest {

    async execute(env: Environment, params: UpdateUserCalendarEventParams, payload: OfficeApiEvent) {
        let service = new ExchangeService(ExchangeVersion.Exchange2010, TimeZoneInfo.Utc);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let rawEvent: OfficeApiEvent = payload;
        let appointment: Appointment;
        if (params.entireSeries) {
            appointment = await Appointment.BindToRecurringMaster(service, new ItemId(params.eventId));
        } else {
            appointment = await Appointment.Bind(service, new ItemId(params.eventId));
        }

        copyApiEventToAppointment(rawEvent, appointment);

        let mode: SendInvitationsOrCancellationsMode = SendInvitationsOrCancellationsMode.SendToNone;

        //If attendees is set & only attendees are updated => Only send update to added attendees
        if (rawEvent.attendees && Object.keys(rawEvent).length === 1) {
            mode = SendInvitationsOrCancellationsMode.SendToChangedAndSaveCopy;
        } else if (rawEvent.attendees && rawEvent.attendees.length > 0 && Object.keys(rawEvent).length > 1) {
            //Otherwise send update to all attendees
            mode = SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy;
        }

        await appointment.Update(ConflictResolutionMode.AutoResolve, mode);
    }
}