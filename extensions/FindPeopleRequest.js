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
var FindPeopleResponse_1 = require('./FindPeopleResponse');

/**
 * @internal Represents a **FindItem** request.
 *
 * @type <TItem>   Item type.
 */
var FindPeopleRequest = (function (_super) {
    __extends(FindPeopleRequest, _super);
    /**
     * @internal Initializes a new instance of the **FindPeopleRequest** class.
     *
     * @param   {ExchangeService}       service             The service.
     * @param   {ServiceErrorHandling}  errorHandlingMode   Indicates how errors should be handled.
     */
    function FindPeopleRequest(service, errorHandlingMode) {
        var _this = _super.call(this, service, errorHandlingMode) || this;
        _this.groupBy = null;
        return _this;
    }

    /**
     * @internal Gets the request version.
     *
     * @return  {ExchangeVersion}      Earliest Exchange version in which this request is supported.
     */
    FindPeopleRequest.prototype.GetMinimumRequiredServerVersion = function () { return ExchangeVersion_1.ExchangeVersion.Exchange2013; };

    /**
     * @internal Gets the name of the response XML element.
     *
     * @return  {string}      XML element name.
     */
    FindPeopleRequest.prototype.GetResponseXmlElementName = function () { return "FindPeopleResponse" };
    /**
     * @internal Gets the name of the XML element.
     *
     * @return  {string} XML element name.
     */
    FindPeopleRequest.prototype.GetXmlElementName = function () { return "FindPeople"; };

    FindPeopleRequest.prototype.ParseResponse = function (jsonBody) {
        var response = new FindPeopleResponse_1.FindPeopleResponse();
        response.LoadFromXmlJsObject(jsonBody, this.Service);
        return response;
    };

    FindPeopleRequest.prototype.WriteElementsToXml = function (writer) {
        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Messages, 'PersonaShape');
        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Types, 'BaseShape', 'IdOnly');
        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Types, 'AdditionalProperties');

        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Types, 'FieldURI');
        writer.WriteAttributeValue('FieldURI', 'persona:DisplayName');
        writer.WriteEndElement();

        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Types, 'FieldURI');
        writer.WriteAttributeValue('FieldURI', 'persona:PersonaType');
        writer.WriteEndElement();

        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Types, 'FieldURI');
        writer.WriteAttributeValue('FieldURI', 'persona:GivenName');
        writer.WriteEndElement();

        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Types, 'FieldURI');
        writer.WriteAttributeValue('FieldURI', 'persona:Surname');
        writer.WriteEndElement();

        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Types, 'FieldURI');
        writer.WriteAttributeValue('FieldURI', 'persona:CompanyName');
        writer.WriteEndElement();

        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Types, 'FieldURI');
        writer.WriteAttributeValue('FieldURI', 'persona:EmailAddress');
        writer.WriteEndElement();

        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Messages, 'IndexedPageItemView');
        writer.WriteAttributeValue('BasePoint', 'Beginning');
        writer.WriteAttributeValue('MaxEntriesReturned', '100');
        writer.WriteAttributeValue('Offset', '0');
        writer.WriteEndElement();

        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Messages, 'ParentFolderId');
        writer.WriteStartElement(XmlNamespace_1.XmlNamespace.Types, 'DistinguishedFolderId');
        writer.WriteAttributeValue('Id', 'directory');
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteElementValue(XmlNamespace_1.XmlNamespace.Messages, 'QueryString', this.QueryString);
    };

    return FindPeopleRequest;
}(FindRequest_1.FindRequest));
exports.FindPeopleRequest = FindPeopleRequest;
