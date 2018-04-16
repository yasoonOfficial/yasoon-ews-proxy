import { Environment } from "../model/proxy";
import { ExchangeService, DeleteMode, ExchangeVersion, FolderId, Folder, Uri, WellKnownFolderName, Mailbox } from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";

export interface DeleteCalendarParams {
    calendarId: string;
    email: string;
}

export class DeleteCalendarRequest {

    async execute(env: Environment, params: DeleteCalendarParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        // Check if
        let targetFolderId = null;
        if (params.calendarId === 'main') {
            targetFolderId = new FolderId(WellKnownFolderName.Calendar, new Mailbox(params.email));
        } else {
            targetFolderId = new FolderId();
            targetFolderId.UniqueId = params.calendarId;
        }

        let targetCalendar: Folder;

        //Bind Calendar to Folder
        try {
            targetCalendar = await Folder.Bind(service, targetFolderId);
        } catch (e) {
            console.log(e.message);
        }

        if (targetCalendar) {
            try {
                targetCalendar.Delete(DeleteMode.HardDelete);
            } catch (e) {
                console.log(e.message);
            }
        }
    }
}