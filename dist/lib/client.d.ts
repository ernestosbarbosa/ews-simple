import * as ews from 'ews-javascript-api';
export declare class ClientBuilder {
    _version: ews.ExchangeVersion;
    _uri: string;
    _user: string | undefined;
    _pwd: string | undefined;
    _token: string | undefined;
    withVersion(value: ews.ExchangeVersion): this;
    withUser(value: string): this;
    withPwd(value: string): this;
    withUri(value: string): this;
    withToken(value: string): this;
    build(): ews.ExchangeService;
}
