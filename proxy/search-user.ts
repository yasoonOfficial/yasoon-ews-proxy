import { ExchangeService, ItemView, Uri } from "ews-javascript-api";
import { applyCredentials } from 'proxy/helper';
import { FindPeopleRequest } from 'extensions/FindPeopleRequest';
import { ProxyMethod, Environment } from 'model/proxy';

export class SearchUserRequest implements ProxyMethod {

    async execute(env: Environment, params: { [key: string]: string }, payload: any) {
        let service = new ExchangeService();
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        var request = new FindPeopleRequest(service, null);
        request.QueryString = params.searchTerm;
        request.View = new ItemView(100);
        let response = await request.Execute();

        return response.People.filter(p => p.PersonaType === 'Person' || p.PersonaType === 'Room').map((p) => ({
            id: p.PersonaId.Id,
            displayName: p.DisplayName,
            mail: p.EmailAddress.EmailAddress,
            givenName: p.GivenName,
            surname: p.Surname,
            relevanceScore: p.RelevanceScore,
            personaType: p.PersonaType
        }));
    }
}