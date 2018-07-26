import { AlternateId, ExchangeService, Folder, IdFormat, NameResolution, ResolveNameSearchLocation, Uri, WellKnownFolderName } from "ews-javascript-api";
import { OfficeUser } from '../model/office';
import { Environment } from "../model/proxy";
import { applyCredentials } from "../proxy/helper";


export class GetOwnUserRequest {

    async execute(env: Environment) {
        let service = new ExchangeService();
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        //If user name !== email address
        let email = env.ewsUser;
        if (!email.includes("@")) {
            //Hack to get own email address
            let inbox = await Folder.Bind(service, WellKnownFolderName.Inbox);
            let aiAlternateid = new AlternateId(IdFormat.EwsId, inbox.Id.UniqueId, "mailbox@domain.com");
            let aiResponse = await service.ConvertId(aiAlternateid, IdFormat.EwsId) as AlternateId;
            email = aiResponse.Mailbox;
        }

        let results = await service.ResolveName(email, ResolveNameSearchLocation.DirectoryOnly, true);
        //Not sure this can really happen... (i mean > 1)
        let res: NameResolution = null;
        if (results.Count == 1) {
            res = results._getItem(0);
        } else {
            let it = results.GetEnumerator();
            it.map((result) => {
                if (result.Mailbox.Address.toLowerCase() === email.toLowerCase()) {
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