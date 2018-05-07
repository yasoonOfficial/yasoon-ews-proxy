import { ExchangeService, ExchangeVersion, Uri } from "ews-javascript-api";
import { CreateGroupRequest as EwsCreateGroupRequest } from '../extensions/CreateGroupRequest';
import { Environment } from '../model/proxy';
import { applyCredentials } from '../proxy/helper';

export class CreateGroupRequest {

    async execute(env: Environment, body: { name: string, alias: string, accessType: string, description: string, autoSubscribeNewMembers: boolean }) {
        let service = new ExchangeService(ExchangeVersion.Exchange2015);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        try {
            //@ts-ignore
            var request = new EwsCreateGroupRequest(service, null);
            //@ts-ignore
            request.Name = body.name;
            //@ts-ignore
            request.Alias = body.alias || body.name.replace(/\W/g, '');
            //@ts-ignore
            request.AccessType = body.accessType || "Public";
            //@ts-ignore
            request.Description = body.description || body.name;
            //@ts-ignore
            request.AutoSubscribeNewMembers = body.autoSubscribeNewMembers || false;
            //@ts-ignore
            let response = await request.Execute();
            let group = response.GroupData;

            return {
                id: group.ExternalDirectoryObjectId,
                name: group.DisplayName,
                mailboxGuid: group.MailboxDatabase,
                smtpAddress: group.GroupIdentity.Value,
                visibility: group.AccessType
            };

        } catch (e) {
            console.log(e.message, e.toString(), e.stack);
            return [];
        }
    }
}