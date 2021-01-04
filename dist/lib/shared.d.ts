import { MessageBody } from "ews-javascript-api";
export interface IEWSConfig {
    uri: string;
    user?: string;
    pwd?: string;
    token?: string;
    to?: string[];
}
export declare enum FolderEnum {
    Inbox = 4,
    Favorites = 43,
    Sent = 8
}
export interface SearchResult {
    subject: string;
    received: string;
    body: string | MessageBody;
    attachments?: {
        [key: string]: string;
    };
}
