import { Environment } from "../model/proxy";
import { ExchangeService, Uri, FolderId, WellKnownFolderName, Mailbox, ExchangeVersion, UserConfigurationProperties } from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { parseString } from "xml2js";
import { XmlCategoriesResult } from "../model/office";

export class GetCategoriesRequest {

    async execute(env: Environment, params: { email: string }) {
        let service = new ExchangeService(ExchangeVersion.Exchange2010);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let calendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(params.email));
        let categoryConfig = await service.GetUserConfiguration("CategoryList", calendar, UserConfigurationProperties.XmlData);
        let rawXml = new Buffer(categoryConfig.XmlData, 'base64').toString();

        let categoriesObject: XmlCategoriesResult = await new Promise<XmlCategoriesResult>((resolve, reject) => {
            parseString(rawXml, (err, result) => {
                if (err)
                    reject(err);
                else
                    resolve(result);
            });
        });

        return categoriesObject.categories.category.map(c => ({
            id: c.$.guid,
            displayName: c.$.name,
            color: 'preset' + c.$.color,
            keyboardShortcut: c.$.keyboardShortcut,
            lastTimeUsed: c.$.lastTimeUsedCalendar,
            usageCount: c.$.usageCount
        }));
    }
}