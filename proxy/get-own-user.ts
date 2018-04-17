import { Environment } from "../model/proxy";
import { ExchangeService, Uri, ResolveNameSearchLocation, NameResolution, Folder, WellKnownFolderName, AlternateId, IdFormat } from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { OfficeUser } from '../model/office';


export class GetOwnUserRequest {

    async execute(env: Environment) {
        let service = new ExchangeService();
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        try {
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

        } catch (e) {
            console.log(e.message, e.toString(), e.stack);
            return [];
        }
    }
}