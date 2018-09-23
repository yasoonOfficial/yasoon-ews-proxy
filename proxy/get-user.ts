import { Environment } from "../model/proxy";
import { ExchangeService, Uri, ResolveNameSearchLocation, NameResolution } from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { OfficeUser } from '../model/office';


export class GetUserRequest {

    async execute(env: Environment, params: { email: string }) {
        let service = new ExchangeService();
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let results = await service.ResolveName(params.email, ResolveNameSearchLocation.DirectoryOnly, true);
        //Not sure this can really happen... (i mean > 1)
        let res: NameResolution = null;
        if (results.Count == 1) {
            res = results._getItem(0);
        } else {
            let it = results.GetEnumerator();
            it.map((result) => {
                if (result.Mailbox.Address.toLowerCase() === params.email.toLowerCase()) {
                    res = result;
                }
            });
        }

        let result: OfficeUser = null
        result = {
            displayName: res.Contact.DisplayName,
            givenName: res.Contact.GivenName,
            id: res.Mailbox.Name,
            mail: res.Mailbox.Address,
            surname: res.Contact.Surname,
            personaType: res.Mailbox.MailboxType.toString()
        }

        return result;
    }
}