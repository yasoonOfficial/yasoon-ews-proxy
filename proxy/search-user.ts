import { ExchangeService, Uri, ResolveNameSearchLocation } from "ews-javascript-api";
import { applyCredentials } from '../proxy/helper';
import { Environment } from '../model/proxy';
import { OfficeUser } from '../model/office';

export class SearchUserRequest {

    async execute(env: Environment, params: { searchTerm: string }) {
        let service = new ExchangeService();
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let resolutions = await service.ResolveName(params.searchTerm, ResolveNameSearchLocation.DirectoryOnly, true);
        let results: OfficeUser[] = [];

        let it = resolutions.GetEnumerator();
        it.map((result) => {
            let officeUser: OfficeUser = {
                displayName: result.Contact.DisplayName,
                givenName: result.Contact.GivenName,
                id: result.Mailbox.Name,
                mail: result.Mailbox.Address,
                surname: result.Contact.Surname,
                personaType: result.Mailbox.MailboxType.toString()
            }
            results.push(officeUser);
        });

        return results;
    }
}