export interface Environment {
    ewsUrl: string;
    ewsAuthType: string;
    ewsToken: string;
    ewsUser: string;
    ewsPassword: string;
}

export interface ProxyMethod {
    execute(env: Environment, params: { [key: string]: string }, payload: any): Promise<any>;
}