import { Environment } from "model/proxy";
import { ExchangeService, Uri, ItemView } from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { FindPeopleRequest } from '../extensions/FindPeopleRequest';

export class GetUserRequest {

    async execute(env: Environment, params: { email: string }) {
        let service = new ExchangeService();
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        var request = new FindPeopleRequest(service, null);
        request.QueryString = params.email;
        request.View = new ItemView(100);
        let response = await request.Execute();

        //Special case for aliases
        if (response.People.length === 1) {
            return response.People.map((p) => ({
                id: p.PersonaId.Id,
                displayName: p.DisplayName,
                mail: p.EmailAddress.EmailAddress,
                givenName: p.GivenName,
                surname: p.Surname,
                personaType: p.PersonaType
            }))[0];
        } else {
            return response.People.filter(p => (p.PersonaType === 'Person' || p.PersonaType === 'Room') && p.EmailAddress.EmailAddress === params.email).map((p) => ({
                id: p.PersonaId.Id,
                displayName: p.DisplayName,
                mail: p.EmailAddress.EmailAddress,
                givenName: p.GivenName,
                surname: p.Surname,
                personaType: p.PersonaType
            }))[0];
        }
    }
}