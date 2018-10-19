import { Environment } from "../model/proxy";
import {
    ExchangeService, Uri, DeleteMode, SendCancellationsMode, ExchangeVersion, ItemId, Appointment, MessageBody
} from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";

export interface DeleteEventParams {
    eventId: string;
    sendCancellations: boolean;
    entireSeries: boolean;
    cancellationMessage: string;
    type: "delete" | "cancel";
    doHardDelete: boolean;
}

export class DeleteEventRequest {

    async execute(env: Environment, params: DeleteEventParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2010);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let sendCancellationsMode = SendCancellationsMode.SendToNone;
        let deleteMode = params.doHardDelete ? DeleteMode.HardDelete : DeleteMode.MoveToDeletedItems;
        if (params.sendCancellations === true)
            sendCancellationsMode = SendCancellationsMode.SendToAllAndSaveCopy;

        let appointment: Appointment;
        if (params.entireSeries === true) {
            appointment = await Appointment.BindToRecurringMaster(service, new ItemId(params.eventId));
        } else {
            appointment = await Appointment.Bind(service, new ItemId(params.eventId));
        }
        if (params.type === "delete") {
            await appointment.Delete(deleteMode, sendCancellationsMode);

        } else if (params.type === "cancel") {
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
                await appointment.Delete(deleteMode, sendCancellationsMode);
            }
        }
    }
}