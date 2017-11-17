import { Environment } from "model/proxy";
import { applyCredentials, validateAutodiscoverRedirection } from "proxy/helper";
import { AutodiscoverService, GetUserSettingsResponse, UserSettingName, ExchangeService, ItemView, Uri } from "ews-javascript-api";
import { AutodiscoverService as NtlmAutodiscoverService } from 'extensions/CustomAutodiscoverService';
import { NtlmExchangeService } from 'extensions/NtlmAutodiscoverService';
import { ntlmAuthXhrApi } from "extensions/CustomNtlmAuthXhrApi";
import { FindPeopleRequest } from 'extensions/FindPeopleRequest';

export class GetAutodiscoverDataRequest {

    async execute(env: Environment, params: { email: string }) {
        let userEmail = params.email;
        let service: ExchangeService = new NtlmExchangeService();
        applyCredentials(service, env);

        await service.AutodiscoverUrl(userEmail, validateAutodiscoverRedirection);

        let discoverService: AutodiscoverService = new NtlmAutodiscoverService();
        discoverService.RedirectionUrlValidationCallback = validateAutodiscoverRedirection;
        applyCredentials(discoverService, env);

        let userSettings: GetUserSettingsResponse;

        userSettings = await discoverService.GetUserSettings(
            userEmail,
            UserSettingName.ExternalWebClientUrls
        );

        let owaUrl = null;
        let ewsUrl = service.Url.AbsoluteUri;
        if (userSettings.Settings && userSettings.Settings[UserSettingName.ExternalWebClientUrls]) {
            let urls: any[] = userSettings.Settings[UserSettingName.ExternalWebClientUrls].Urls;
            if (urls && urls.length > 0) {
                owaUrl = urls[0].Url;
            }
        }

        let authMode = 'ntlm';
        //Try if we should use ntlm or basic, fallback to ntlm
        try {
            //To-do move this to own function
            let testService = new ExchangeService();
            testService.Url = new Uri(ewsUrl);
            testService.XHRApi = new ntlmAuthXhrApi(userEmail, new Buffer(env.ewsPassword, 'base64').toString());
            testService.UseDefaultCredentials = true; //Bug... 

            var request = new FindPeopleRequest(testService, null);
            request.QueryString = userEmail;
            request.View = new ItemView(100);
            await request.Execute();
        } catch (e) {
            authMode = 'basic';
        }

        let type = (ewsUrl === 'https://outlook.office365.com/EWS/Exchange.asmx') ? 'office365' : 'onpremise';
        return { type: type, url: ewsUrl, owaUrl: owaUrl, authMode: authMode };
    }
}