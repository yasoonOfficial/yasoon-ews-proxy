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
var ExchangeVersion_1 = require("ews-javascript-api/js/Enumerations/ExchangeVersion");
var FindItemResponse_1 = require("ews-javascript-api/js/Core/Responses/FindItemResponse");
var XmlElementNames_1 = require("ews-javascript-api/js/Core/XmlElementNames");
var SimpleServiceRequestBase_1 = require("ews-javascript-api/js/Core/Requests/SimpleServiceRequestBase");
var XmlNamespace_1 = require("ews-javascript-api/js/Enumerations/XmlNamespace");
var GetUserPhotoResponse_1 = require('./GetUserPhotoResponse');

/**
 * @internal Represents a **FindItem** request.
 *
 * @type <TItem>   Item type.
 */
var GetUserPhotoRequest = (function (_super) {
    __extends(GetUserPhotoRequest, _super);
    /**
     * @internal Initializes a new instance of the **GetUserPhotoRequest** class.
     *
     * @param   {ExchangeService}       service             The service.
     * @param   {ServiceErrorHandling}  errorHandlingMode   Indicates how errors should be handled.
     */
    function GetUserPhotoRequest(service, errorHandlingMode) {
        var _this = _super.call(this, service, errorHandlingMode) || this;
        _this.EmailAddress = null;
        _this.Size = 128;
        return _this;
    }

    /**
     * @internal Gets the request version.
     *
     * @return  {ExchangeVersion}      Earliest Exchange version in which this request is supported.
     */
    GetUserPhotoRequest.prototype.GetMinimumRequiredServerVersion = function () { return ExchangeVersion_1.ExchangeVersion.Exchange2013; };

    /**
     * @internal Gets the name of the response XML element.
     *
     * @return  {string}      XML element name.
     */
    GetUserPhotoRequest.prototype.GetResponseXmlElementName = function () { return "GetUserPhotoResponse" };
    /**
     * @internal Gets the name of the XML element.
     *
     * @return  {string} XML element name.
     */
    GetUserPhotoRequest.prototype.GetXmlElementName = function () { return "GetUserPhoto"; };

    /**
     * @internal Executes this request.
     *
     * @return  {Promise<GetRoomsResponse>}      Service response  :Promise.
     */
    GetUserPhotoRequest.prototype.Execute = function () {
        return this.InternalExecute().then(function (serviceResponse) {
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        });
    };

    GetUserPhotoRequest.prototype.ParseResponse = function (jsonBody) {
        var response = new GetUserPhotoResponse_1.GetUserPhotoResponse();
        response.LoadFromXmlJsObject(jsonBody, this.Service);
        return response;
    };

    GetUserPhotoRequest.prototype.WriteElementsToXml = function (writer) {
        //<m:Email>sadie@contoso.com</m:Email >
        //<m:SizeRequested>HR360x360</m:SizeRequested >
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'Email', this.EmailAddress);
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'SizeRequested', 'HR' + this.Size + "x" + this.Size);
    };

    return GetUserPhotoRequest;
}(SimpleServiceRequestBase_1.SimpleServiceRequestBase));
exports.GetUserPhotoRequest = GetUserPhotoRequest;
