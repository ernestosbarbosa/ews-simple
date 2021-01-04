import { WellKnownFolderName, MessageBody } from "ews-javascript-api";

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
