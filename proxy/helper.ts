import { WebCredentials, OAuthCredentials, SoapFaultDetails, EffectiveRights } from "ews-javascript-api";
import { ntlmAuthXhrApi } from "../extensions/CustomNtlmAuthXhrApi";
import { Environment } from "../model/proxy";
import { EWS_AUTH_TYPE_HEADER, EWS_PASSWORD_HEADER, EWS_TOKEN_HEADER, EWS_URL_HEADER, EWS_USER_HEADER, EWS_URL_OFFICE_365, PROXY_SECRET_HEADER } from "../model/constants";

import * as express from 'express';

import { createTableService, TableUtilities } from 'azure-storage';


export function applyCredentials(service: any, env: Environment) {
    if (env.ewsToken) {
        service.Credentials = new OAuthCredentials(env.ewsToken);
    } else if (env.ewsUser && env.ewsPassword && env.ewsAuthType === 'ntlm') {
        let userEmail = env.ewsUser;
        let password;
        if (env.ewsPassword.indexOf('$') > 0) {
            password = {
                lmHash: new Buffer(env.ewsPassword.split('$')[0], 'base64'),
                ntlmHash: new Buffer(env.ewsPassword.split('$')[1], 'base64')
            };
        } else {
            password = new Buffer(env.ewsPassword, 'base64').toString();
        }

        service.XHRApi = new ntlmAuthXhrApi(userEmail, password, true);
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

export function decodeUrlId(id: string) {
    return id.replace(new RegExp("[-]", "g"), "/").replace(new RegExp("[_]", "g"), "+");
}

export function getEnvFromHeader(req: express.Request, secret: string): Environment {
    if (secret && req.headers[PROXY_SECRET_HEADER] !== secret) {
        throw new Error('Invalid Secret');
    }

    let env = {
        ewsAuthType: req.headers[EWS_AUTH_TYPE_HEADER],
        ewsToken: req.headers[EWS_TOKEN_HEADER],
        ewsUrl: req.headers[EWS_URL_HEADER] || EWS_URL_OFFICE_365,
        ewsUser: req.headers[EWS_USER_HEADER],
        ewsPassword: req.headers[EWS_PASSWORD_HEADER],
        tableService: undefined,
        logId: new Date().getTime(),
        logCount: 0
    };

    if (process.env.LOGGING_CONNECTION_STRING) {
        try {
            env.tableService = createTableService(process.env.LOGGING_CONNECTION_STRING);
        } catch (e) {
        }
    } else {
        env.tableService = {
            insertEntity: () => null
        }
    }

    return env;
}

export function getAccessArrayFromEffectiveRights(effectiveRights: any) {
    if (effectiveRights && Number.isInteger(effectiveRights)) {
        let access = [];
        if ((effectiveRights & EffectiveRights.CreateContents) !== 0)
            access.push('create');
        if ((effectiveRights & EffectiveRights.CreateHierarchy) !== 0)
            access.push('createFolder');
        if ((effectiveRights & EffectiveRights.Delete) !== 0)
            access.push('delete');
        if ((effectiveRights & EffectiveRights.Modify) !== 0)
            access.push('edit');
        if ((effectiveRights & EffectiveRights.Read) !== 0)
            access.push('read');

        return access;
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

export function calculateStoreId(userDn: string, serverName: string) {
    let userDnHex = ''
    for (let i = 0; i < userDn.length; i++)
        userDnHex += (userDn.charCodeAt(i) >>> 0).toString(16).toUpperCase();

    let serverNameHex = '';
    for (let i = 0; i < serverName.length; i++)
        serverNameHex += (serverName.charCodeAt(i) >>> 0).toString(16).toUpperCase();

    let flags = "00000000";
    let ProviderUID = "38A1BB1005E5101AA1BB08002B2A56C2";
    let versionFlag = "0000";
    let DLLFileName = "454D534D44422E444C4C00000000";
    let WrappedFlags = "00000000";
    let WrappedProviderUID = "1B55FA20AA6611CD9BC800AA002FC45A";
    let WrappedType = "0C000000";
    let StoredIdStringHex = flags + ProviderUID + versionFlag + DLLFileName + WrappedFlags + WrappedProviderUID + WrappedType + serverNameHex + "00" + userDnHex + "00";
    return StoredIdStringHex;
    /*
    let sender = "confluence@yasoon.com";
    let something = '';
    for (let i = 0; i < sender.length; i++)
        something = (sender.charCodeAt(i) >>> 0).toString(16).toUpperCase() + '00';
 
    return StoredIdStringHex + 'E94632F4480000000200000010000000' + something + '00000000';*/
}

export function calculateAddressBookId(userDn: string) {
    let userDnHex = ''
    for (let i = 0; i < userDn.length; i++)
        userDnHex += (userDn.charCodeAt(i) >>> 0).toString(16).toUpperCase();

    let Provider = "00000000DCA740C8C042101AB4B908002B2FE1820100000000000000";
    let userdnStringHex = Provider + userDnHex + "00";
    return userdnStringHex;
}

export function guidToBytes(guid) {
    var bytes = [];
    guid.split('-').map((number, index) => {
        var bytesInChar = index < 3 ? number.match(/.{1,2}/g).reverse() : number.match(/.{1,2}/g);
        bytesInChar.map((byte) => { bytes.push(parseInt(byte, 16)); })
    });
    return bytes;
}

export function validateAutodiscoverRedirection(redirectionUrl: string) {
    //Todo
    //return redirectionUrl === 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc';
    return true;
}

export function requestWrapper(secret: string, func: (req: express.Request, res: express.Response, env: Environment) => Promise<any>): (req: express.Request, res: express.Response) => void {
    return (async (req: express.Request, res: express.Response) => {
        let env = getEnvFromHeader(req, secret);
        try {
            if (env.tableService) {
                let gen = TableUtilities.entityGenerator;
                env.tableService.insertEntity('ewsevents', {
                    PartitionKey: gen.String(env.ewsUser || req.params.email || "unknown"),
                    RowKey: gen.String(env.logId + "-" + env.logCount++),
                    path: gen.String(req.path),
                    key: gen.String('request-init'),
                    ewsUrl: gen.String(env.ewsUrl),
                    authType: gen.String(env.ewsAuthType)
                }, function () {
                });
            }

            await func(req, res, env);
        }
        catch (e) {
            let message = '';
            if (e instanceof SoapFaultDetails) {
                message = e.Message;
                res.status(e.HttpStatusCode).send(e.Message);
            } else if (e.message) {
                message = e.message;
                res.status(500).send(e.message);
            }
            else if (typeof e === 'string') {
                message = e;
                res.status(500).send(e);
            } else {
                res.status(500).send();
            }

            if (env.tableService) {
                let gen = TableUtilities.entityGenerator;
                env.tableService.insertEntity('ewsevents', {
                    PartitionKey: gen.String(env.ewsUser || req.params.email || "unknown"),
                    RowKey: gen.String(env.logId + "-" + env.logCount++),
                    path: gen.String(req.path),
                    key: gen.String('request-error'),
                    error: gen.String(message),
                    stack: gen.String(e.stack || "")
                }, () => null);
            }

            //Valid response, don't log anything
            if (e.HttpStatusCode === 404)
                return;

            if (e && e.message) {
                console.log(e.message);
            }
            if (e && e.stack) {
                console.log(e.stack);
            }
            if (e && e.toString) {
                console.log(e.toString());
            }
        }
    });
}

export function getCredentialsAsAuth(env: Environment) {
    if (env.ewsToken) {
        return 'Bearer ' + env.ewsToken;
    }
    else if (env.ewsUser && env.ewsPassword) {
        let password = new Buffer(env.ewsPassword, 'base64').toString();
        return 'Basic ' + Buffer.from(env.ewsUser + ':' + password).toString('base64');
    }
}