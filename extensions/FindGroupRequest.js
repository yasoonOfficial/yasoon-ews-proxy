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
var FindRequest_1 = require("ews-javascript-api/js/Core/Requests/FindRequest");
var XmlNamespace_1 = require("ews-javascript-api/js/Enumerations/XmlNamespace");
var FindGroupResponse_1 = require('./FindGroupResponse');

/**
 * @internal Represents a **FindItem** request.
 *
 * @type <TItem>   Item type.
 */
var FindGroupRequest = (function (_super) {
    __extends(FindGroupRequest, _super);
    /**
     * @internal Initializes a new instance of the **FindGroupRequest** class.
     *
     * @param   {ExchangeService}       service             The service.
     * @param   {ServiceErrorHandling}  errorHandlingMode   Indicates how errors should be handled.
     */
    function FindGroupRequest(service, errorHandlingMode) {
        var _this = _super.call(this, service, errorHandlingMode) || this;
        _this.groupBy = null;
        return _this;
    }

    /**
     * @internal Gets the request version.
     *
     * @return  {ExchangeVersion}      Earliest Exchange version in which this request is supported.
     */
    FindGroupRequest.prototype.GetMinimumRequiredServerVersion = function () { return "V2016_07_13"; };

    /**
     * @internal Gets the name of the response XML element.
     *
     * @return  {string}      XML element name.
     */
    FindGroupRequest.prototype.GetResponseXmlElementName = function () { return "FindUnifiedGroupsResponseMessage" };
    /**
     * @internal Gets the name of the XML element.
     *
     * @return  {string} XML element name.
     */
    FindGroupRequest.prototype.GetXmlElementName = function () { return "FindUnifiedGroups"; };

    FindGroupRequest.prototype.ParseResponse = function (jsonBody) {
        var response = new FindGroupResponse_1.FindGroupResponse();
        response.LoadFromXmlJsObject(jsonBody, this.Service);
        return response;
    };

    FindGroupRequest.prototype.WriteElementsToXml = function (writer) {
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'QueryString', this.QueryString);
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'PageSize', "100");
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'IncludeInactiveGroups', "true");
    };

    return FindGroupRequest;
}(FindRequest_1.FindRequest));
exports.FindGroupRequest = FindGroupRequest;
