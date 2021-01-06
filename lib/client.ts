/*
 * EWS client builder
 *
 * 
 */

// Dependencies
import * as ews from 'ews-javascript-api';
import * as util from 'util';

import { ClientVersion } from './shared';

const debug = util.debuglog('client');

export class ClientBuilder {
  _version = ews.ExchangeVersion.Exchange2013;
  _uri = 'https://outlook.office365.com/Ews/Exchange.asmx';
  _user: string | undefined;
  _pwd: string | undefined;
  _token: string | undefined;

  withVersion(value: ClientVersion) {
    switch (value) {
      case ClientVersion.Exchange2010:
        this._version = ews.ExchangeVersion.Exchange2010;
        break;
      case ClientVersion.Exchange2015:
        this._version = ews.ExchangeVersion.Exchange2015;
        break;
      case ClientVersion.Exchange2016:
        this._version = ews.ExchangeVersion.Exchange2016;
        break;
      case ClientVersion.Exchange2013:
      default:
        this._version = ews.ExchangeVersion.Exchange2013;
    }
    return this;
  }

  withUser(value: string) {
    this._user = value;
    return this;
  }

  withPwd(value: string) {
    this._pwd = value;
    return this;
  }

  withUri(value: string) {
    this._uri = value;
    return this;
  }

  withToken(value: string) {
    this._token = value;
    return this;
  }

  build() {
    debug("Configuring service...");
    const service = new ews.ExchangeService(this._version);
    if (!!this._token) {
      service.Credentials = new ews.OAuthCredentials(this._token);
    } else if (!!this._user && !!this._pwd) {
      service.Credentials = new ews.WebCredentials(this._user, this._pwd);
    } else {
      throw new Error(`Please specify user and pwd or token to continue!`);
    }
    service.Url = new ews.Uri(this._uri);
    debug("Configuration done");
    return service;
  }
}
