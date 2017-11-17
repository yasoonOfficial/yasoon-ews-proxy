import { WebCredentials, ExchangeServiceBase, OAuthCredentials } from "ews-javascript-api";
import { ntlmAuthXhrApi } from "extensions/CustomNtlmAuthXhrApi";
import { Environment } from "model/proxy";
import { EWS_AUTH_TYPE_HEADER, EWS_PASSWORD_HEADER, EWS_TOKEN_HEADER, EWS_URL_HEADER, EWS_USER_HEADER } from "model/constants";

import * as express from 'express';

export function applyCredentials(service: ExchangeServiceBase, env: Environment) {
    if (env.ewsToken) {
        service.Credentials = new OAuthCredentials(env.ewsToken);
    } else if (env.ewsUser && env.ewsPassword && env.ewsAuthType === 'ntlm') {
        let userEmail = env.ewsUser;
        let password = new Buffer(env.ewsPassword, 'base64').toString();
        service.XHRApi = new ntlmAuthXhrApi(userEmail, password);
        service.UseDefaultCredentials = true; //Bug... 
    } else if (env.ewsUser && env.ewsPassword) {
        let userEmail = env.ewsUser;
        let password = new Buffer(env.ewsPassword, 'base64').toString();
        service.Credentials = new WebCredentials(userEmail, password);
    }
    else {
        throw new Error('No Auth!');
    }
}

export function getEnvFromHeader(req: express.Request): Environment {
    return {
        ewsAuthType: req.headers[EWS_AUTH_TYPE_HEADER],
        ewsToken: req.headers[EWS_TOKEN_HEADER],
        ewsUrl: req.headers[EWS_URL_HEADER],
        ewsUser: req.headers[EWS_USER_HEADER],
        ewsPassword: req.headers[EWS_PASSWORD_HEADER]
    };
}

export function getAccessArrayFromEffectiveRights(effectiveRights: any) {
    if (effectiveRights && Number.isInteger(effectiveRights)) {
        //Todo!!!
        throw new Error('not supported yet, todo!');
        //Bug in ews-javascript-api, see https://github.com/gautamsi/ews-javascript-api/pull/214
    } else if (effectiveRights && effectiveRights['__type'] === 'EffectiveRights') {
        let rights: {
            CreateAssociated: "true" | "false",
            CreateContents: "true" | "false",
            CreateHierarchy: "true" | "false",
            Delete: "true" | "false",
            Modify: "true" | "false",
            Read: "true" | "false",
            ViewPrivateItems: "true" | "false"
        } = <any>effectiveRights;

        let access = [];
        if (rights.CreateContents === 'true')
            access.push('create');
        if (rights.CreateHierarchy === 'true')
            access.push('createFolder');
        if (rights.Delete === 'true')
            access.push('delete');
        if (rights.Modify === 'true')
            access.push('edit');
        if (rights.Read === 'true')
            access.push('read');

        return access;
    } else {
        return [];
    }
}