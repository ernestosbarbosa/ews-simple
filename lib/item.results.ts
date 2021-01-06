/*
 * EWS FindItemsResult builder
 *
 * 
 */

// Dependencies
import * as ews from 'ews-javascript-api';
import * as util from 'util';
import * as fs from 'fs';
import * as path from 'path';

const debug = util.debuglog('results');

import { FolderEnum, SearchResult } from './shared';

export class FindItemsResultBuilder {
  _service: ews.ExchangeService | undefined;
  _folder = ews.WellKnownFolderName.Inbox;
  _basePropertySet = ews.BasePropertySet.FirstClassProperties;
  _looupPropertySet = new ews.PropertySet(this._basePropertySet);
  _takeN = 1000;
  _offset: number = 0;
  _offsetBasePoint: ews.OffsetBasePoint = ews.OffsetBasePoint.Beginning;
  _querystring = 'isread:true';
  _readPropertySet = new ews.PropertySet(this._basePropertySet, [ews.ItemSchema.Attachments, ews.ItemSchema.HasAttachments, ews.EmailMessageSchema.Body]);
  _markAsRead = true;
  _delete = false;
  _saveAttachments = false;
  _getRaw = false;
  _rawBody = false;
  _downloadFolder = '../downloads';

  withService(value: ews.ExchangeService) {
    this._service = value;
    return this;
  }

  withFolder(value: FolderEnum) {
    switch (value) {
      case FolderEnum.Favorites:
        this._folder = ews.WellKnownFolderName.Favorites;
        break;
      case FolderEnum.Sent:
        this._folder = ews.WellKnownFolderName.SentItems;
        break;
      case FolderEnum.Inbox:
      default:
        this._folder = ews.WellKnownFolderName.Inbox;
    }
    return this;
  }

  withTakeN(value: number) {
    this._takeN = value;
    return this;
  }

  withOffset(value: number) {
    this._offset = value;
    return this;
  }

  withOffsetBasePoint(value: 'start' | 'end') {
    this._offsetBasePoint = value === 'start' ? ews.OffsetBasePoint.Beginning : ews.OffsetBasePoint.End;
    return this;
  }

  withQueryString(value: string) {
    this._querystring = value;
    return this;
  }

  withMarkAsRead(value: boolean) {
    this._markAsRead = value;
    return this;
  }

  withDelete(value: boolean) {
    this._delete = value;
    return this;
  }

  withGetRaw(value: boolean) {
    this._getRaw = value;
    return this;
  }

  withSaveAttachments(value: boolean) {
    this._saveAttachments = value;
    return this;
  }

  withDownloadsFolder(value: string) {
    this._downloadFolder = value;
    return this;
  }

  withRawBody(value: boolean) {
    this._rawBody = value;
    return this;
  }

  async getItems() {
    if (this._service === undefined) {
      throw new Error(`Ews service wasn't initialized!`);
    }
    debug(`Bind folder`);
    const mailInbox = await ews.Folder.Bind(this._service, this._folder);

    debug(`Init view`);
    const view = new ews.ItemView(this._takeN, this._offset, this._offsetBasePoint);
    view.PropertySet = this._looupPropertySet;

    debug(`Find items...`);
    const result = await mailInbox.FindItems(this._querystring, view);
    debug(`Find total count: ${result.TotalCount}`);

    return result.Items;
  }

  async execute() {
    const out: SearchResult[] = [];
    const items = await this.getItems();
    debug(`Total Items found: ${items.length} `);

    if (this._getRaw) {
      return items;
    }

    debug(`--------------Start Processing--------------`);
    for (const item of items) {
      if (item instanceof ews.EmailMessage) {
        debug(`Loading properties...`);
        await item.Load(this._readPropertySet);

        const itemObject: SearchResult = {
          subject: item.Subject,
          received: item.DateTimeReceived.toString(),
          body: this._rawBody ? item.Body : item.Body.Text.replace(/<\/?[^>]+(>|$)/g, '')
        };

        for (const attachment of item.Attachments.GetEnumerator()) {
          if (attachment instanceof ews.FileAttachment) {
            debug(`Load attachment content...`);
            await attachment.Load();
            if (itemObject.attachments === undefined) {
              itemObject.attachments = {};
            }
            itemObject.attachments[attachment.Name] = Buffer.from(attachment.Base64Content, 'base64').toString('utf-8');

            if (this._saveAttachments) {
              debug(`Saving attachment content...`);
              try {
                await fs.writeFileSync(path.join(__dirname, this._downloadFolder, attachment.Name), itemObject.attachments[attachment.Name]);
              } catch (err) {
                console.error(err);
              }
            }
          }
        }

        if (this._markAsRead) {
          debug(`Mark as read and save state`);
          item.IsRead = true;
          /*
           * await item.Update(conflictResolution);
           *  where conflictResolution has next numeric values:
           *    - AlwaysOverwrite:  2  Local property changes overwrite server-side changes.
           *    - AutoResolve:      1  Local property changes are applied to the server unless the server-side copy is more recent than the local copy.
           *    - NeverOverwrite:   0  Local property changes are discarded.
           */
          await item.Update(1);
        }

        if (this._delete) {
          debug(`Delete if specified`);
          /*
           * await item.delete(deleteMode, suppressReadReceipts);
           *  where deleteMode has next numeric values
           *    - HardDelete:         0 The item or folder will be permanently deleted.
           *    - MoveToDeletedItems: 2	The item or folder will be moved to the mailbox's Deleted Items folder.
           *    - SoftDelete:         1 The item or folder will be moved to the dumpster. Items and folders in the dumpster can be recovered.
           *  where suppressReadReceipts Boolean- true if read receipts should not be sent if the item being deleted has requested a read receipt; otherwise, false.
           */
          await item.Delete(2, true);
        }
        out.push(itemObject);
      }
    }
    debug(`--------------Stop Processing--------------`);
    return out;
  }
}
