import { Environment } from "model/proxy";
import {
    ExchangeService, Uri, DeleteMode, SendCancellationsMode, AffectedTaskOccurrence, ExchangeVersion, ItemId, Appointment
} from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";

export interface DeleteEventParams {
    eventId: string;
    sendCancellationsMode: SendCancellationsMode
    affectedTaskOccurrence: AffectedTaskOccurrence;
    type: "delete" | "decline";
}

export class DeleteEventRequest {

    async execute(env: Environment, params: DeleteEventParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        //if (params.type === "delete") {
        let itemId = new ItemId(params.eventId);
        await service.DeleteItems([itemId], DeleteMode.MoveToDeletedItems, params.sendCancellationsMode, params.affectedTaskOccurrence);
        /*
                } else if (params.type === "decline"){
                    let appointment = await Appointment.Bind(service, new ItemId(params.eventId));
                    appointment.Decline();
                }
        */


    }
}