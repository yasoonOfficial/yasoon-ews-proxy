import { AutodiscoverService, ExchangeService, ExchangeServiceBase, ExchangeVersion, GetUserSettingsResponse, ResolveNameSearchLocation, Uri, UserSettingName, WebCredentials } from "ews-javascript-api";
import { AutodiscoverService as CustomAutodiscoverService } from '../extensions/CustomAutodiscoverService';
import { ntlmAuthXhrApi } from "../extensions/CustomNtlmAuthXhrApi";
import { Environment } from "../model/proxy";
import { validateAutodiscoverRedirection } from "../proxy/helper";
import { isNullOrEmpty } from './mapper';
import { TableUtilities } from "azure-storage";

export class GetAutodiscoverDataRequest {

    async execute(env: Environment, params: { email: string }) {
        let userEmail = params.email || env.ewsUser;
        let userName = env.ewsUser;
        let password = new Buffer(env.ewsPassword, 'base64').toString();

        let discoverService: AutodiscoverService = <any>new CustomAutodiscoverService();
        discoverService.RedirectionUrlValidationCallback = validateAutodiscoverRedirection;

        let userSettings: GetUserSettingsResponse;

        //First try basic with userName & pw, then NTLM with userName & pw
        let credentials = [
            {
                authMode: 'basic',
                userNameRequired: false,
                apply: (svc: ExchangeServiceBase) => svc.Credentials = new WebCredentials(userEmail, password)
            },
            {
                authMode: 'basic',
                userNameRequired: true,
                apply: (svc: ExchangeServiceBase) => svc.Credentials = new WebCredentials(userName, password)
            },
            {
                authMode: 'ntlm',
                userNameRequired: false,
                apply: (svc: ExchangeServiceBase) => {
                    svc.Credentials = null;
                    svc.XHRApi = new ntlmAuthXhrApi(userEmail, password, true);
                    svc.UseDefaultCredentials = true; //Bug... 
                }
            },
            {
                authMode: 'ntlm',
                userNameRequired: true,
                apply: (svc: ExchangeServiceBase) => {
                    svc.Credentials = null;
                    svc.XHRApi = new ntlmAuthXhrApi(userName, password, true);
                    svc.UseDefaultCredentials = true; //Bug... 
                }
            }
        ];

        let errors = [];
        let gen = TableUtilities.entityGenerator;

        for (const credential of credentials) {
            try {
                //Logging...
                env.tableService.insertEntity('ewsevents', {
                    PartitionKey: gen.String(userEmail),
                    RowKey: gen.String(env.logId + "-" + env.logCount++),
                    key: gen.String('autodiscover-ews-url'),
                    data: gen.String(JSON.stringify({ authMode: credential.authMode, userNameRequired: credential.userNameRequired }))
                }, () => null);

                credential.apply(discoverService);
                userSettings = await discoverService.GetUserSettings(
                    userEmail,
                    UserSettingName.InternalEwsUrl,
                    UserSettingName.ExternalEwsUrl,
                    UserSettingName.ExternalMailboxServer,
                    UserSettingName.EwsSupportedSchemas
                );

                break;
            }
            catch (e) {
                env.tableService.insertEntity('ewsevents', {
                    PartitionKey: gen.String(userEmail),
                    RowKey: gen.String(env.logId + "-" + env.logCount++),
                    key: gen.String('autodiscover-ews-url'),
                    data: gen.String(JSON.stringify({ authMode: credential.authMode, userNameRequired: credential.userNameRequired })),
                    error: gen.String(e.toString())
                }, () => null);

                errors.push(e.toString());
            }
        }

        //If we don't have user settings by now, abort..
        if (!userSettings) {
            return {
                success: false,
                errorMessage: "Couldn't connect to autodiscover service... \r\n" + errors.join('\r\n\r\n')
            };
        }

        //Determine correct ews url
        let ewsUrl = userSettings.Settings[UserSettingName.ExternalEwsUrl];
        if (!discoverService.IsExternal || isNullOrEmpty(ewsUrl)) {
            let intUrl = userSettings.Settings[UserSettingName.InternalEwsUrl];
            if (!isNullOrEmpty(intUrl)) {
                ewsUrl = intUrl;
            }
        }

        if (!ewsUrl) {
            return {
                success: false,
                errorMessage: "Couldn't retrieve ews URL... \r\n" + errors.join('\r\n\r\n')
            };
        }

        //Use resolve service to find working configuration..
        let exchangeService = new ExchangeService(ExchangeVersion.Exchange2010);
        exchangeService.Url = new Uri(ewsUrl);

        let authMode = null;
        let userNameRequired: boolean = null;

        for (const credential of credentials) {
            try {
                env.tableService.insertEntity('ewsevents', {
                    PartitionKey: gen.String(userEmail),
                    RowKey: gen.String(env.logId + "-" + env.logCount++),
                    key: gen.String('autodiscover-resolve-user'),
                    data: gen.String(JSON.stringify({ authMode: credential.authMode, userNameRequired: credential.userNameRequired }))
                }, () => null);

                credential.apply(exchangeService);
                await exchangeService.ResolveName(params.email, ResolveNameSearchLocation.DirectoryOnly, true);

                //It worked, take over parameters & break
                authMode = credential.authMode;
                userNameRequired = credential.userNameRequired;

                break;
            }
            catch (e) {
                env.tableService.insertEntity('ewsevents', {
                    PartitionKey: gen.String(userEmail),
                    RowKey: gen.String(env.logId + "-" + env.logCount++),
                    key: gen.String('autodiscover-resolve-user'),
                    data: gen.String(JSON.stringify({ authMode: credential.authMode, userNameRequired: credential.userNameRequired })),
                    error: gen.String(e.toString())
                }, () => null);

                errors.push(e.toString());
            }
        }

        if (authMode === null || userNameRequired === null) {
            return {
                success: false,
                errorMessage: "Couldn't connect to resolve-service \r\n" + errors.join('\r\n\r\n')
            };
        }

        let extHost: string = userSettings.Settings[UserSettingName.ExternalMailboxServer];
        let ewsSupportedSchemas: string = userSettings.Settings[UserSettingName.EwsSupportedSchemas];

        let mode = "unknown";
        if ("outlook.office365.com" === extHost) {
            mode = "office365";
        } else if ("partner.outlook.cn" === extHost) {
            mode = "office365:china";
        } else if ("outlook.office.de" === extHost) {
            mode = "office365:germany";
        } else if (extHost != null && extHost.includes("office365.us")) {
            mode = "office365:gov";
        } else if (ewsSupportedSchemas.includes("2013")) {
            mode = "onpremise";
        } else if (ewsSupportedSchemas.includes("2010")) {
            mode = "onpremise2010";
        }

        env.tableService.insertEntity('ewsevents', {
            PartitionKey: gen.String(userEmail),
            RowKey: gen.String(env.logId + "-" + env.logCount++),
            key: gen.String('autodiscover-detected-settings'),
            data: gen.String(JSON.stringify({
                ewsUrl: ewsUrl,
                mode: mode,
                authMode: authMode,
                ewsSupportedSchemas: ewsSupportedSchemas
            }))
        }, () => null);

        return { success: true, mode: mode, url: ewsUrl, authMode: authMode, externalMailboxServer: extHost, ewsSupportedSchemas: ewsSupportedSchemas, userNameRequired: userNameRequired };
    }
}
