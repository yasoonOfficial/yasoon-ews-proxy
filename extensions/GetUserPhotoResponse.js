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
var XmlElementNames_1 = require("../node_modules/ews-javascript-api/js/Core/XmlElementNames");
var XmlAttributeNames_1 = require("../node_modules/ews-javascript-api/js/Core/XmlAttributeNames");
var ServiceResponse_1 = require("../node_modules/ews-javascript-api/js/Core/Responses/ServiceResponse");
var EwsServiceJsonReader_1 = require("../node_modules/ews-javascript-api/js/Core/EwsServiceJsonReader");

var GetUserPhotoResponse = (function (_super) {
    __extends(GetUserPhotoResponse, _super);
    function GetUserPhotoResponse(isGrouped, propertySet) {
        var _this = _super.call(this) || this;
        _this.PictureData = null;
        return _this;
    }

    /**
     * @internal Reads response elements from Xml JsObject.
     *
     * @param   {any}               jsObject   The response object.
     * @param   {ExchangeService}   service    The service.
     */
    GetUserPhotoResponse.prototype.ReadElementsFromXmlJsObject = function (responseObject, service) {
        this.PictureData = responseObject;
    };

    return GetUserPhotoResponse;
}(ServiceResponse_1.ServiceResponse));
exports.GetUserPhotoResponse = GetUserPhotoResponse;
