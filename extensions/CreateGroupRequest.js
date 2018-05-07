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
var CreateGroupResponse_1 = require('./CreateGroupResponse');

/**
 * @internal Represents a **FindItem** request.
 *
 * @type <TItem>   Item type.
 */
var CreateGroupRequest = (function (_super) {
    __extends(CreateGroupRequest, _super);
    /**
     * @internal Initializes a new instance of the **CreateGroupRequest** class.
     *
     * @param   {ExchangeService}       service             The service.
     * @param   {ServiceErrorHandling}  errorHandlingMode   Indicates how errors should be handled.
     */
    function CreateGroupRequest(service, errorHandlingMode) {
        var _this = _super.call(this, service, errorHandlingMode) || this;
        _this.Name = null;
        _this.Alias = null;
        _this.AccessType = "Public";
        _this.Description = null;
        _this.AutoSubscribeNewMembers = false;
        return _this;
    }

    /**
     * @internal Gets the request version.
     *
     * @return  {ExchangeVersion}      Earliest Exchange version in which this request is supported.
     */
    CreateGroupRequest.prototype.GetMinimumRequiredServerVersion = function () { return ExchangeVersion_1.ExchangeVersion.Exchange2015; };

    /**
     * @internal Gets the name of the response XML element.
     *
     * @return  {string}      XML element name.
     */
    CreateGroupRequest.prototype.GetResponseXmlElementName = function () { return "CreateUnifiedGroupResponseMessage" };
    /**
     * @internal Gets the name of the XML element.
     *
     * @return  {string} XML element name.
     */
    CreateGroupRequest.prototype.GetXmlElementName = function () { return "CreateUnifiedGroup"; };

    /**
     * @internal Executes this request.
     *
     * @return  {Promise<GetRoomsResponse>}      Service response  :Promise.
     */
    CreateGroupRequest.prototype.Execute = function () {
        return this.InternalExecute().then(function (serviceResponse) {
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        });
    };

    CreateGroupRequest.prototype.ParseResponse = function (jsonBody) {
        var response = new CreateGroupResponse_1.CreateGroupResponse();
        response.LoadFromXmlJsObject(jsonBody, this.Service);
        return response;
    };

    CreateGroupRequest.prototype.WriteElementsToXml = function (writer) {
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'Name', this.Name);
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'Alias', this.Alias);
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Types, 'AccessType', this.AccessType);
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'Description', this.Description);
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'AutoSubscribeNewMembers', this.AutoSubscribeNewMembers);
    };

    return CreateGroupRequest;
}(SimpleServiceRequestBase_1.SimpleServiceRequestBase));
exports.CreateGroupRequest = CreateGroupRequest;
