import { WellKnownFolderName, MessageBody, ExchangeVersion } from "ews-javascript-api";

export interface IEWSConfig {
  uri: string;
  user?: string;
  pwd?: string;
  token?: string;
  to?: string[];
};

export enum FolderEnum {
  Inbox = WellKnownFolderName.Inbox,
  Favorites = WellKnownFolderName.Favorites,
  Sent = WellKnownFolderName.SentItems
};

export interface SearchResult {
  subject: string;
  received: string;
  body: string | MessageBody;
  attachments?: { [key: string]: string }
};

export enum ClientVersion {
  Exchange2010 = ExchangeVersion.Exchange2010,
  Exchange2013 = ExchangeVersion.Exchange2013,
  Exchange2015 = ExchangeVersion.Exchange2015,
  Exchange2016 = ExchangeVersion.Exchange2016
}