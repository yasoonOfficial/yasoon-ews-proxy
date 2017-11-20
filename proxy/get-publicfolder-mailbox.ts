import { Environment } from "model/proxy";
import { applyCredentials, getCredentialsAsAuth, validateAutodiscoverRedirection } from "../proxy/helper";
import { AutodiscoverService, GetUserSettingsResponse, UserSettingName } from "ews-javascript-api";
import { AutodiscoverService as NtlmAutodiscoverService } from '../extensions/CustomAutodiscoverService';
import * as poxAutodiscover from 'autodiscover';

export class GetPublicFolderMailboxRequest {

    async execute(env: Environment, params: { email: string }) {
        let userEmail = params.email;
        let service: AutodiscoverService = new NtlmAutodiscoverService();
        service.RedirectionUrlValidationCallback = validateAutodiscoverRedirection;
        applyCredentials(service, env);

        let userSettings: GetUserSettingsResponse;

        userSettings = await service.GetUserSettings(
            userEmail,
            UserSettingName.PublicFolderInformation
        );

        let anchorMailbox = userSettings.Settings[UserSettingName.PublicFolderInformation];

        //No public folder access?
        if (!anchorMailbox) {
            return { success: false };
        }

        return new Promise((resolve, reject) => {
            poxAutodiscover.getPOXAutodiscoverValues(anchorMailbox, getCredentialsAsAuth(env), (err, data) => {
                if (!err && data && data.Autodiscover && data.Autodiscover.Response.length > 0) {
                    let server = data.Autodiscover.Response[0].Account[0].Protocol.find(p => p.Type[0] === 'EXCH').Server[0];
                    resolve({ success: true, anchorMailbox: userSettings.Settings[UserSettingName.PublicFolderInformation], publicFolderMailbox: server });
                } else {
                    resolve({ success: false, err: err.toString() });
                }
            });
        });
    }
}