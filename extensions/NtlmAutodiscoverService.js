"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var ExchangeService_1 = require("../node_modules/ews-javascript-api/js/Core/ExchangeService");
var AutodiscoverLocalException_1 = require("../node_modules/ews-javascript-api/js/Exceptions/AutodiscoverLocalException");
var AutodiscoverService_1 = require("./CustomAutodiscoverService");
var AutodiscoverErrorCode_1 = require("../node_modules/ews-javascript-api/js/Enumerations/AutodiscoverErrorCode");
var ServiceErrorHandling_1 = require("../node_modules/ews-javascript-api/js/Enumerations/ServiceErrorHandling");
var ServiceLocalException_1 = require("../node_modules/ews-javascript-api/js/Exceptions/ServiceLocalException");
var ServiceRemoteException_1 = require("../node_modules/ews-javascript-api/js/Exceptions/ServiceRemoteException");
var Strings_1 = require("../node_modules/ews-javascript-api/js/Strings");
var ExtensionMethods_1 = require("../node_modules/ews-javascript-api/js/ExtensionMethods");
var UserSettingName_1 = require("../node_modules/ews-javascript-api/js/Enumerations/UserSettingName");

/**
 * Represents a binding to the **Exchange Web Services**.
 *
 */
var NtlmAutodiscoverService = (function (_super) {
    __extends(NtlmAutodiscoverService, _super);

    function NtlmAutodiscoverService() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
    }

    NtlmAutodiscoverService.prototype.GetAutodiscoverUrl = function (emailAddress, requestedServerVersion, validateRedirectionUrlCallback) {
        var _this = this;
        var autodiscoverService = new AutodiscoverService_1.AutodiscoverService(null, null, requestedServerVersion);
        autodiscoverService.Credentials = this.Credentials;
        autodiscoverService.XHRApi = this.XHRApi;
        autodiscoverService.RedirectionUrlValidationCallback = validateRedirectionUrlCallback,
            autodiscoverService.EnableScpLookup = this.EnableScpLookup;
        return autodiscoverService.GetUserSettings(emailAddress, UserSettingName_1.UserSettingName.InternalEwsUrl, UserSettingName_1.UserSettingName.ExternalEwsUrl)
            .then(function (response) {
                switch (response.ErrorCode) {
                    case AutodiscoverErrorCode_1.AutodiscoverErrorCode.NoError:
                        return _this.GetEwsUrlFromResponse(response, autodiscoverService.IsExternal);
                    case AutodiscoverErrorCode_1.AutodiscoverErrorCode.InvalidUser:
                        throw new ServiceRemoteException_1.ServiceRemoteException(ExtensionMethods_1.StringHelper.Format(Strings_1.Strings.InvalidUser, emailAddress));
                    case AutodiscoverErrorCode_1.AutodiscoverErrorCode.InvalidRequest:
                        throw new ServiceRemoteException_1.ServiceRemoteException(ExtensionMethods_1.StringHelper.Format(Strings_1.Strings.InvalidAutodiscoverRequest, response.ErrorMessage));
                    default:
                        _this.TraceMessage(TraceFlags_1.TraceFlags.AutodiscoverConfiguration, ExtensionMethods_1.StringHelper.Format("No EWS Url returned for user {0}, error code is {1}", emailAddress, response.ErrorCode));
                        throw new ServiceRemoteException_1.ServiceRemoteException(response.ErrorMessage);
                }
            }, function (err) {
                throw err;
            });
    };

    return NtlmAutodiscoverService;
}(ExchangeService_1.ExchangeService));

exports.NtlmExchangeService = NtlmAutodiscoverService;
