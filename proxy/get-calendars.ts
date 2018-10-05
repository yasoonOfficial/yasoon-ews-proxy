import { Environment } from "../model/proxy";
import { ExchangeService, Uri, FolderId, WellKnownFolderName, Mailbox, FolderView, BasePropertySet, PropertySet, FolderSchema, ExchangeVersion } from "ews-javascript-api";
import { applyCredentials, getAccessArrayFromEffectiveRights } from "../proxy/helper";

export class GetCalendarsRequest {

    async execute(env: Environment, params: { email: string }) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let sharedCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(params.email));
        let folderView = new FolderView(1000);
        folderView.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.EffectiveRights);

        try {
        let ewsResult = await service.FindFolders(sharedCalendar, folderView);
        return ewsResult.Folders.map(f => ({
            id: f.Id.UniqueId,
            name: f.DisplayName,
            access: getAccessArrayFromEffectiveRights(f.EffectiveRights)
        }));
    }
        catch (e) {
            console.log(e.message, e.toString(), e.stack);
            return [];
        }
    }
}