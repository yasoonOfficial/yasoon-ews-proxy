import { ExchangeService, Uri, ItemView } from "ews-javascript-api";
import { applyCredentials } from '../proxy/helper';
import { Environment } from '../model/proxy';
import { FindGroupRequest as EwsFindGroupRequest } from '../extensions/FindGroupRequest';

export class FindGroupRequest {

    async execute(env: Environment, params: { searchTerm: string }) {
        let service = new ExchangeService();
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        try {
            //@ts-ignore
            var request = new EwsFindGroupRequest(service, null);
            //@ts-ignore
            request.QueryString = params.searchTerm;
            //@ts-ignore
            request.View = new ItemView(100);

            //@ts-ignore
            let response = await request.Execute();
            return response.groups.map(g => ({
                type: g['__type'],
                id: g.ExternalDirectoryObjectId,
                name: g.DisplayName,
                description: g.AdditionalProperties.Description,
                mailboxGuid: g.MailboxGuid,
                smtpAddress: g.SmtpAddress,
                visibility: g.AccessType,
                isFavorite: g.IsFavorite
            }));

        } catch (e) {
            console.log(e.message, e.toString(), e.stack);
            return [];
        }
    }
}