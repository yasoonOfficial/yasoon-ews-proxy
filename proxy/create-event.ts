import { Environment } from "../model/proxy";
import {
    ExchangeService, Uri, FolderId, WellKnownFolderName,
    Mailbox, Appointment, ExchangeVersion, TimeZoneInfo,
    SendInvitationsMode
} from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { copyApiEventToAppointment } from '../proxy/mapper';
import { OfficeApiEvent } from '../model/office';

export interface CreateEventParams {
    email: string;
    calendarId: string;
}

export class CreateEventRequest {

    async execute(env: Environment, params: CreateEventParams, payload: OfficeApiEvent) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let rawEvent: OfficeApiEvent = payload;
        if (!rawEvent.subject || !rawEvent.start || !rawEvent.end)
            throw new Error('Missing data');

        let targetFolderId: FolderId = null;
        let mode: SendInvitationsMode = SendInvitationsMode.SendToNone;

        if (rawEvent.attendees && rawEvent.attendees.length > 0) {
            mode = SendInvitationsMode.SendToAllAndSaveCopy;
        }

        //Get folder instance
        if (params.calendarId === 'main') {
            targetFolderId = new FolderId(WellKnownFolderName.Calendar, new Mailbox(params.email));
        } else {
            targetFolderId = new FolderId();
            targetFolderId.UniqueId = params.calendarId;
        }

        let appointment = new Appointment(service);
        copyApiEventToAppointment(rawEvent, appointment);

        await appointment.Save(targetFolderId, mode);
        return { id: appointment.Id.UniqueId };
    }
}