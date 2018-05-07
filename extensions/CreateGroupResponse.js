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
var XmlElementNames_1 = require("ews-javascript-api/js/Core/XmlElementNames");
var XmlAttributeNames_1 = require("ews-javascript-api/js/Core/XmlAttributeNames");
var ServiceResponse_1 = require("ews-javascript-api/js/Core/Responses/ServiceResponse");
var EwsServiceJsonReader_1 = require("ews-javascript-api/js/Core/EwsServiceJsonReader");

var CreateGroupResponse = (function (_super) {
    __extends(CreateGroupResponse, _super);
    function CreateGroupResponse(isGrouped, propertySet) {
        var _this = _super.call(this) || this;
        _this.GroupData = null;
        return _this;
    }

    /**
     * @internal Reads response elements from Xml JsObject.
     *
     * @param   {any}               jsObject   The response object.
     * @param   {ExchangeService}   service    The service.
     */
    CreateGroupResponse.prototype.ReadElementsFromXmlJsObject = function (responseObject, service) {
        this.GroupData = responseObject;
    };

    return CreateGroupResponse;
}(ServiceResponse_1.ServiceResponse));
exports.CreateGroupResponse = CreateGroupResponse;
