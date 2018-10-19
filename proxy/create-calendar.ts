import { Environment } from "../model/proxy";
import {
    ExchangeService, Uri, FolderId, WellKnownFolderName,
    Mailbox, ExchangeVersion, CalendarFolder,
    FolderPermission, StandardUser, FolderPermissionLevel
} from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";

export class CreateCalendarRequest {

    async execute(env: Environment, params: { email: string }, calendarData: { name: string, isOwnMailbox: boolean }) {
        let service = new ExchangeService(ExchangeVersion.Exchange2010);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let folder = new CalendarFolder(service);
        folder.DisplayName = calendarData.name;

        let parentFolder = new FolderId(WellKnownFolderName.Calendar, new Mailbox(params.email));
        await folder.Save(parentFolder);

        //Yeah.. Well.. Exchange
        if (calendarData.isOwnMailbox) {
            return new Promise((resolve, reject) => {
                setTimeout(async () => {
                    let defaultEdit = new FolderPermission(StandardUser.Default, FolderPermissionLevel.Editor);
                    folder.Permissions.Add(defaultEdit);
                    await folder.Update();

                    resolve({
                        id: folder.Id.UniqueId,
                        name: folder.DisplayName
                    });
                }, 5000);
            });
        } else {
            return {
                id: folder.Id.UniqueId,
                name: folder.DisplayName
            };
        }
    }
}