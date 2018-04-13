import { Environment } from "../model/proxy";
import { applyCredentials, validateAutodiscoverRedirection } from "../proxy/helper";
import { AutodiscoverService, GetUserSettingsResponse, UserSettingName, ExchangeService, ItemView, Uri, WebCredentials } from "ews-javascript-api";
import { AutodiscoverService as NtlmAutodiscoverService } from '../extensions/CustomAutodiscoverService';
import { NtlmExchangeService } from '../extensions/NtlmAutodiscoverService';
import { ntlmAuthXhrApi } from "../extensions/CustomNtlmAuthXhrApi";
import { FindPeopleRequest } from '../extensions/FindPeopleRequest';
import { isNullOrEmpty } from './mapper';

export class GetAutodiscoverDataRequest {

    async execute(env: Environment, params: { email: string }) {
        let userEmail = params.email;
        let userName = env.ewsUser;
        let service: ExchangeService = <any>new NtlmExchangeService();
        applyCredentials(service, env);

        await service.AutodiscoverUrl(userEmail, validateAutodiscoverRedirection);

        let discoverService: AutodiscoverService = <any>new NtlmAutodiscoverService();
        discoverService.RedirectionUrlValidationCallback = validateAutodiscoverRedirection;
        applyCredentials(discoverService, env);

        let userSettings: GetUserSettingsResponse;

        userSettings = await discoverService.GetUserSettings(
            userEmail,
            UserSettingName.ExternalMailboxServer,
            UserSettingName.EwsSupportedSchemas
        );

        let ewsUrl = service.Url.AbsoluteUri;
        let authMode = 'ntlm';
        let userNameRequired: boolean = false;
        let errorMessage = '';
        let success: boolean = false;

        //Try if we should use ntlm or basic, fallback to ntlm
        try {
            //To-do move this to own function
            //Check if ntlm works with User Email
            let testService = new ExchangeService();
            testService.Url = new Uri(ewsUrl);
            testService.XHRApi = new ntlmAuthXhrApi(userEmail, new Buffer(env.ewsPassword, 'base64').toString());
            testService.UseDefaultCredentials = true; //Bug... 

            var request = <any>new FindPeopleRequest(testService, null);
            request.QueryString = userEmail;
            request.View = new ItemView(100);
            await request.Execute();
            success = true;
        } catch (e) {
            errorMessage += "Error Message from Email + PW / NTLM";
            try {
                //If not, check ntlm with User Name
                let testService = new ExchangeService();
                testService.Url = new Uri(ewsUrl);
                testService.XHRApi = new ntlmAuthXhrApi(userName, new Buffer(env.ewsPassword, 'base64').toString());
                testService.UseDefaultCredentials = true; //Bug... 

                var request = <any>new FindPeopleRequest(testService, null);
                request.QueryString = userEmail;
                request.View = new ItemView(100);
                await request.Execute();
                userNameRequired = true;
                success = true;
            } catch (e) {
                errorMessage += "Error Message from User + PW / NTLM";
                try {
                    let testService = new ExchangeService();
                    testService.Credentials = new WebCredentials(userEmail, env.ewsPassword);
                    authMode = 'basic';
                    success = true;
                } catch (e) {
                    errorMessage += "Error Message from Email + PW / Basic";
                }
            }
        }

        if (isNullOrEmpty(ewsUrl)) {
            errorMessage += "No internal or external EWS Url could be found";
            success = false;
        }

        let extHost = userSettings.Settings[UserSettingName.ExternalMailboxServer];
        let ewsSupportedSchemas = userSettings.Settings[UserSettingName.EwsSupportedSchemas];

        let mode = "unknown";
        if ("outlook.office365.com" === extHost) {
            mode = "office365:public";
        } else if ("partner.outlook.cn" === extHost) {
            mode = "office365:china";
        } else if ("outlook.office.de" === extHost) {
            mode = "office365:germany";
        } else if (extHost != null && extHost.contains("office365.us")) {
            mode = "office365:gov";
        } else if (ewsSupportedSchemas.contains("2013")) {
            mode = "onpremise2013";
        } else if (ewsSupportedSchemas.contains("2010")) {
            mode = "onpremise2010";
        }

        return { success: success, errorMessage: errorMessage, mode: mode, url: ewsUrl, authMode: authMode, extHost: extHost, ewsSupportedSchemas: ewsSupportedSchemas, userNameRequired: userNameRequired };
    }
}
