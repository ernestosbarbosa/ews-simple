"use strict";
/*
 * EWS FindItemsResult builder
 *
 *
 */
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    Object.defineProperty(o, k2, { enumerable: true, get: function() { return m[k]; } });
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.FindItemsResultBuilder = void 0;
// Dependencies
var ews = __importStar(require("ews-javascript-api"));
var util = __importStar(require("util"));
var fs = __importStar(require("fs"));
var path = __importStar(require("path"));
var debug = util.debuglog('results');
var shared_1 = require("./shared");
var FindItemsResultBuilder = /** @class */ (function () {
    function FindItemsResultBuilder() {
        this._folder = ews.WellKnownFolderName.Inbox;
        this._basePropertySet = ews.BasePropertySet.FirstClassProperties;
        this._looupPropertySet = new ews.PropertySet(this._basePropertySet);
        this._takeN = 1000;
        this._offset = 0;
        this._offsetBasePoint = ews.OffsetBasePoint.Beginning;
        this._querystring = 'isread:true';
        this._readPropertySet = new ews.PropertySet(this._basePropertySet, [ews.ItemSchema.Attachments, ews.ItemSchema.HasAttachments, ews.EmailMessageSchema.Body]);
        this._markAsRead = true;
        this._delete = false;
        this._saveAttachments = false;
        this._getRaw = false;
        this._rawBody = false;
        this._downloadFolder = '../downloads';
    }
    FindItemsResultBuilder.prototype.withService = function (value) {
        this._service = value;
        return this;
    };
    FindItemsResultBuilder.prototype.withFolder = function (value) {
        switch (value) {
            case shared_1.FolderEnum.Favorites:
                this._folder = ews.WellKnownFolderName.Favorites;
                break;
            case shared_1.FolderEnum.Sent:
                this._folder = ews.WellKnownFolderName.SentItems;
                break;
            case shared_1.FolderEnum.Inbox:
            default:
                this._folder = ews.WellKnownFolderName.Inbox;
        }
        return this;
    };
    FindItemsResultBuilder.prototype.withTakeN = function (value) {
        this._takeN = value;
        return this;
    };
    FindItemsResultBuilder.prototype.withOffset = function (value) {
        this._offset = value;
        return this;
    };
    FindItemsResultBuilder.prototype.withOffsetBasePoint = function (value) {
        this._offsetBasePoint = value === 'start' ? ews.OffsetBasePoint.Beginning : ews.OffsetBasePoint.End;
        return this;
    };
    FindItemsResultBuilder.prototype.withQueryString = function (value) {
        this._querystring = value;
        return this;
    };
    FindItemsResultBuilder.prototype.withMarkAsRead = function (value) {
        this._markAsRead = value;
        return this;
    };
    FindItemsResultBuilder.prototype.withDelete = function (value) {
        this._delete = value;
        return this;
    };
    FindItemsResultBuilder.prototype.withGetRaw = function (value) {
        this._getRaw = value;
        return this;
    };
    FindItemsResultBuilder.prototype.withSaveAttachments = function (value) {
        this._saveAttachments = value;
        return this;
    };
    FindItemsResultBuilder.prototype.withDownloadsFolder = function (value) {
        this._downloadFolder = value;
        return this;
    };
    FindItemsResultBuilder.prototype.withRawBody = function (value) {
        this._rawBody = value;
        return this;
    };
    FindItemsResultBuilder.prototype.getItems = function () {
        return __awaiter(this, void 0, void 0, function () {
            var mailInbox, view, result;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this._service === undefined) {
                            throw new Error("Ews service wasn't initialized!");
                        }
                        debug("Bind folder");
                        return [4 /*yield*/, ews.Folder.Bind(this._service, this._folder)];
                    case 1:
                        mailInbox = _a.sent();
                        debug("Init view");
                        view = new ews.ItemView(this._takeN, this._offset, this._offsetBasePoint);
                        view.PropertySet = this._looupPropertySet;
                        debug("Find items...");
                        return [4 /*yield*/, mailInbox.FindItems(this._querystring, view)];
                    case 2:
                        result = _a.sent();
                        debug("Find total count: " + result.TotalCount);
                        return [2 /*return*/, result.Items];
                }
            });
        });
    };
    FindItemsResultBuilder.prototype.execute = function () {
        return __awaiter(this, void 0, void 0, function () {
            var out, items, _i, items_1, item, itemObject, _a, _b, attachment, err_1;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        out = [];
                        return [4 /*yield*/, this.getItems()];
                    case 1:
                        items = _c.sent();
                        debug("Total Items found: " + items.length + " ");
                        if (this._getRaw) {
                            return [2 /*return*/, items];
                        }
                        debug("--------------Start Processing--------------");
                        _i = 0, items_1 = items;
                        _c.label = 2;
                    case 2:
                        if (!(_i < items_1.length)) return [3 /*break*/, 16];
                        item = items_1[_i];
                        if (!(item instanceof ews.EmailMessage)) return [3 /*break*/, 15];
                        debug("Loading properties...");
                        return [4 /*yield*/, item.Load(this._readPropertySet)];
                    case 3:
                        _c.sent();
                        itemObject = {
                            subject: item.Subject,
                            received: item.DateTimeReceived.toString(),
                            body: this._rawBody ? item.Body : item.Body.Text.replace(/<\/?[^>]+(>|$)/g, '')
                        };
                        _a = 0, _b = item.Attachments.GetEnumerator();
                        _c.label = 4;
                    case 4:
                        if (!(_a < _b.length)) return [3 /*break*/, 10];
                        attachment = _b[_a];
                        if (!(attachment instanceof ews.FileAttachment)) return [3 /*break*/, 9];
                        debug("Load attachment content...");
                        return [4 /*yield*/, attachment.Load()];
                    case 5:
                        _c.sent();
                        if (itemObject.attachments === undefined) {
                            itemObject.attachments = {};
                        }
                        itemObject.attachments[attachment.Name] = Buffer.from(attachment.Base64Content, 'base64').toString('utf-8');
                        if (!this._saveAttachments) return [3 /*break*/, 9];
                        debug("Saving attachment content...");
                        _c.label = 6;
                    case 6:
                        _c.trys.push([6, 8, , 9]);
                        return [4 /*yield*/, fs.writeFileSync(path.join(__dirname, this._downloadFolder, attachment.Name), itemObject.attachments[attachment.Name])];
                    case 7:
                        _c.sent();
                        return [3 /*break*/, 9];
                    case 8:
                        err_1 = _c.sent();
                        console.error(err_1);
                        return [3 /*break*/, 9];
                    case 9:
                        _a++;
                        return [3 /*break*/, 4];
                    case 10:
                        if (!this._markAsRead) return [3 /*break*/, 12];
                        debug("Mark as read and save state");
                        item.IsRead = true;
                        /*
                         * await item.Update(conflictResolution);
                         *  where conflictResolution has next numeric values:
                         *    - AlwaysOverwrite:  2  Local property changes overwrite server-side changes.
                         *    - AutoResolve:      1  Local property changes are applied to the server unless the server-side copy is more recent than the local copy.
                         *    - NeverOverwrite:   0  Local property changes are discarded.
                         */
                        return [4 /*yield*/, item.Update(1)];
                    case 11:
                        /*
                         * await item.Update(conflictResolution);
                         *  where conflictResolution has next numeric values:
                         *    - AlwaysOverwrite:  2  Local property changes overwrite server-side changes.
                         *    - AutoResolve:      1  Local property changes are applied to the server unless the server-side copy is more recent than the local copy.
                         *    - NeverOverwrite:   0  Local property changes are discarded.
                         */
                        _c.sent();
                        _c.label = 12;
                    case 12:
                        if (!this._delete) return [3 /*break*/, 14];
                        debug("Delete if specified");
                        /*
                         * await item.delete(deleteMode, suppressReadReceipts);
                         *  where deleteMode has next numeric values
                         *    - HardDelete:         0 The item or folder will be permanently deleted.
                         *    - MoveToDeletedItems: 2	The item or folder will be moved to the mailbox's Deleted Items folder.
                         *    - SoftDelete:         1 The item or folder will be moved to the dumpster. Items and folders in the dumpster can be recovered.
                         *  where suppressReadReceipts Boolean- true if read receipts should not be sent if the item being deleted has requested a read receipt; otherwise, false.
                         */
                        return [4 /*yield*/, item.Delete(2, true)];
                    case 13:
                        /*
                         * await item.delete(deleteMode, suppressReadReceipts);
                         *  where deleteMode has next numeric values
                         *    - HardDelete:         0 The item or folder will be permanently deleted.
                         *    - MoveToDeletedItems: 2	The item or folder will be moved to the mailbox's Deleted Items folder.
                         *    - SoftDelete:         1 The item or folder will be moved to the dumpster. Items and folders in the dumpster can be recovered.
                         *  where suppressReadReceipts Boolean- true if read receipts should not be sent if the item being deleted has requested a read receipt; otherwise, false.
                         */
                        _c.sent();
                        _c.label = 14;
                    case 14:
                        out.push(itemObject);
                        _c.label = 15;
                    case 15:
                        _i++;
                        return [3 /*break*/, 2];
                    case 16:
                        debug("--------------Stop Processing--------------");
                        return [2 /*return*/, out];
                }
            });
        });
    };
    return FindItemsResultBuilder;
}());
exports.FindItemsResultBuilder = FindItemsResultBuilder;
