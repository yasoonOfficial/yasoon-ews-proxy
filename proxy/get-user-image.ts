import { Environment } from "../model/proxy";
import { ExchangeService, Uri } from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { GetUserPhotoRequest } from '../extensions/GetUserPhotoRequest';

export class GetUserImageRequest {

    async execute(env: Environment, params: { email: string }) {
        let service = new ExchangeService();
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        //@ts-ignore
        var request = new GetUserPhotoRequest(service, null);
        //@ts-ignore
        request.EmailAddress = params.email;
        //@ts-ignore
        request.Size = 360;

        let response = await request.Execute();
        return {
            mimeType: response.PictureData.ContentType,
            content: new Buffer(response.PictureData.PictureData, 'base64')
        };
    }
}