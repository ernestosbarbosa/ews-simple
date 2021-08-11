"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ClientVersion = exports.FolderEnum = void 0;
;
var FolderEnum;
(function (FolderEnum) {
    FolderEnum[FolderEnum["Inbox"] = 4] = "Inbox";
    FolderEnum[FolderEnum["Favorites"] = 43] = "Favorites";
    FolderEnum[FolderEnum["Sent"] = 8] = "Sent";
})(FolderEnum = exports.FolderEnum || (exports.FolderEnum = {}));
;
;
var ClientVersion;
(function (ClientVersion) {
    ClientVersion[ClientVersion["Exchange2010"] = 1] = "Exchange2010";
    ClientVersion[ClientVersion["Exchange2013"] = 4] = "Exchange2013";
    ClientVersion[ClientVersion["Exchange2015"] = 6] = "Exchange2015";
    ClientVersion[ClientVersion["Exchange2016"] = 7] = "Exchange2016";
})(ClientVersion = exports.ClientVersion || (exports.ClientVersion = {}));
