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

var FindGroupResponse = (function (_super) {
    __extends(FindGroupResponse, _super);
    function FindGroupResponse(isGrouped, propertySet) {
        var _this = _super.call(this) || this;
        _this.groups = [];
        return _this;
    }

    Object.defineProperty(FindGroupResponse.prototype, "Groups", {
        /**
         * Gets collection for all rooms returned
         */
        get: function () {
            return this.groups;
        },
        enumerable: true,
        configurable: true
    });

    /**
     * @internal Reads response elements from Xml JsObject.
     *
     * @param   {any}               jsObject   The response object.
     * @param   {ExchangeService}   service    The service.
     */
    FindGroupResponse.prototype.ReadElementsFromXmlJsObject = function (responseObject, service) {
        var groups = responseObject["GroupsSets"]["UnifiedGroupsSet"]["Groups"];
        var responses = EwsServiceJsonReader_1.EwsServiceJsonReader.ReadAsArray(groups, "UnifiedGroup");
        for (var _i = 0, responses_1 = responses; _i < responses_1.length; _i++) {
            this.groups.push(responses_1[_i]);
        }
    };

    return FindGroupResponse;
}(ServiceResponse_1.ServiceResponse));
exports.FindGroupResponse = FindGroupResponse;
