import { Environment } from "../model/proxy";
import { applyCredentials, validateAutodiscoverRedirection } from "../proxy/helper";
import { AutodiscoverService, GetUserSettingsResponse, UserSettingName, ExchangeService, ItemView, Uri } from "ews-javascript-api";
import { AutodiscoverService as NtlmAutodiscoverService } from '../extensions/CustomAutodiscoverService';
import { NtlmExchangeService } from '../extensions/NtlmAutodiscoverService';
import { ntlmAuthXhrApi } from "../extensions/CustomNtlmAuthXhrApi";
import { FindPeopleRequest } from '../extensions/FindPeopleRequest';

export class GetAutodiscoverDataRequest {

    async execute(env: Environment, params: { email: string, user: string }) {
        let userEmail = params.email;
        let userName = params.user;
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
        } catch (e) {
            try {
                //If not, check ntlm with userName
                let testService = new ExchangeService();
                testService.Url = new Uri(ewsUrl);
                testService.XHRApi = new ntlmAuthXhrApi(userName, new Buffer(env.ewsPassword, 'base64').toString());
                testService.UseDefaultCredentials = true; //Bug... 

                var request = <any>new FindPeopleRequest(testService, null);
                request.QueryString = userEmail;
                request.View = new ItemView(100);
                await request.Execute();
                userNameRequired = true;
            } catch (e) {
                authMode = 'basic';
            }
        }

        let externalMailboxServer = userSettings.Settings[UserSettingName.ExternalMailboxServer];
        let ewsSupportedSchemas = userSettings.Settings[UserSettingName.EwsSupportedSchemas];
        let type = 'notSupported';

        if (ewsUrl === 'https://outlook.office365.com/EWS/Exchange.asmx') {
            type = 'office365';
        } else if (ewsSupportedSchemas.includes("2013")) {
            type = 'onpremise2013';
        } else if (ewsSupportedSchemas.includes("2010")) {
            type = "onpremise2010";
        }

        return { type: type, url: ewsUrl, authMode: authMode, externalMailboxServer: externalMailboxServer, ewsSupportedSchemas: ewsSupportedSchemas, userNameRequired: userNameRequired };
    }
}