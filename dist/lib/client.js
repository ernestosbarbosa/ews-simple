"use strict";
/*
 * EWS client builder
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.ClientBuilder = void 0;
// Dependencies
var ews = __importStar(require("ews-javascript-api"));
var util = __importStar(require("util"));
var shared_1 = require("./shared");
var debug = util.debuglog('client');
var ClientBuilder = /** @class */ (function () {
    function ClientBuilder() {
        this._version = ews.ExchangeVersion.Exchange2013;
        this._uri = 'https://outlook.office365.com/Ews/Exchange.asmx';
    }
    ClientBuilder.prototype.withVersion = function (value) {
        switch (value) {
            case shared_1.ClientVersion.Exchange2010:
                this._version = ews.ExchangeVersion.Exchange2010;
                break;
            case shared_1.ClientVersion.Exchange2015:
                this._version = ews.ExchangeVersion.Exchange2015;
                break;
            case shared_1.ClientVersion.Exchange2016:
                this._version = ews.ExchangeVersion.Exchange2016;
                break;
            case shared_1.ClientVersion.Exchange2013:
            default:
                this._version = ews.ExchangeVersion.Exchange2013;
        }
        return this;
    };
    ClientBuilder.prototype.withUser = function (value) {
        this._user = value;
        return this;
    };
    ClientBuilder.prototype.withPwd = function (value) {
        this._pwd = value;
        return this;
    };
    ClientBuilder.prototype.withUri = function (value) {
        this._uri = value;
        return this;
    };
    ClientBuilder.prototype.withToken = function (value) {
        this._token = value;
        return this;
    };
    ClientBuilder.prototype.build = function () {
        debug("Configuring service...");
        var service = new ews.ExchangeService(this._version);
        if (!!this._token) {
            service.Credentials = new ews.OAuthCredentials(this._token);
        }
        else if (!!this._user && !!this._pwd) {
            service.Credentials = new ews.WebCredentials(this._user, this._pwd);
        }
        else {
            throw new Error("Please specify user and pwd or token to continue!");
        }
        service.Url = new ews.Uri(this._uri);
        debug("Configuration done");
        return service;
    };
    return ClientBuilder;
}());
exports.ClientBuilder = ClientBuilder;
