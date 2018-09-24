import { ExchangeService, Uri, ResolveNameSearchLocation } from "ews-javascript-api";
import { applyCredentials } from '../proxy/helper';
import { Environment } from '../model/proxy';

export class SearchUserRequest {

    async execute(env: Environment, params: { searchTerm: string }) {
        let service = new ExchangeService();
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let resolutions = await service.ResolveName(params.searchTerm, ResolveNameSearchLocation.DirectoryOnly, true);
        return resolutions.GetEnumerator().map((result) => ({
            displayName: result.Contact.DisplayName,
            givenName: result.Contact.GivenName,
            id: result.Mailbox.Name,
            mail: result.Mailbox.Address,
            surname: result.Contact.Surname,
            personaType: result.Mailbox.MailboxType.toString()
        }));
    }
}