import { Environment } from "model/proxy";
import {
    ExchangeService, Uri, DeleteMode, SendCancellationsMode, AffectedTaskOccurrence, ExchangeVersion, ItemId, Appointment, MessageBody
} from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";

export interface DeleteEventParams {
    eventId: string;
    sendCancellations: boolean;
    entireSeries: boolean;
    cancellationMessage: string;
    type: "delete" | "cancel";
}

export class DeleteEventRequest {

    async execute(env: Environment, params: DeleteEventParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let sendCancellationsMode = SendCancellationsMode.SendToNone;
        if (params.sendCancellations === true)
            sendCancellationsMode = SendCancellationsMode.SendToAllAndSaveCopy;

        let affectedTaskOccurrence = AffectedTaskOccurrence.SpecifiedOccurrenceOnly;
        if (params.entireSeries === true)
            affectedTaskOccurrence = AffectedTaskOccurrence.AllOccurrences;

        if (params.type === "delete") {
            let itemId = new ItemId(params.eventId);
            await service.DeleteItems([itemId], DeleteMode.MoveToDeletedItems, sendCancellationsMode, affectedTaskOccurrence);

        } else if (params.type === "cancel") {
            let appointment: Appointment;
            if (affectedTaskOccurrence === AffectedTaskOccurrence.AllOccurrences) {
                appointment = await Appointment.BindToRecurringMaster(service, new ItemId(params.eventId));
            } else if (affectedTaskOccurrence === AffectedTaskOccurrence.SpecifiedOccurrenceOnly) {
                appointment = await Appointment.Bind(service, new ItemId(params.eventId));
            }

            if (appointment.IsMeeting && sendCancellationsMode === SendCancellationsMode.SendToAllAndSaveCopy) {
                // If it's a meeting and the user is the organizer, cancel it.
                // Determine this by testing the AppointmentState bitmask for 
                // the presence of the second bit. This bit indicates that the appointment
                // was received, which means that someone sent it to the user. Therefore,
                // they're not the organizer.
                let isReceived = 2;
                if ((appointment.AppointmentState & isReceived) == 0) {
                    await appointment.CancelMeeting(params.cancellationMessage);
                }
                // If it's a meeting and the user is not the organizer, decline it.
                else {
                    if (params.cancellationMessage) {
                        let declineMessage = appointment.CreateDeclineMessage();
                        declineMessage.Body = new MessageBody(params.cancellationMessage);
                        //Todo: Fix API
                        declineMessage.Sensitivity = <any>"Private";
                        await declineMessage.Send();
                    } else {
                        await appointment.Decline(true);
                    }
                }
            }
            else {
                // The item isn't a meeting, so just delete it.
                await appointment.Delete(DeleteMode.MoveToDeletedItems, sendCancellationsMode);
            }
        }
    }
}