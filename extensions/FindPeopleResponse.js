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

var FindPeopleResponse = (function (_super) {
    __extends(FindPeopleResponse, _super);
    function FindPeopleResponse(isGrouped, propertySet) {
        var _this = _super.call(this) || this;
        _this.people = [];
        return _this;
    }

    Object.defineProperty(FindPeopleResponse.prototype, "People", {
        /**
         * Gets collection for all rooms returned
         */
        get: function () {
            return this.people;
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
    FindPeopleResponse.prototype.ReadElementsFromXmlJsObject = function (responseObject, service) {
        var responses = EwsServiceJsonReader_1.EwsServiceJsonReader.ReadAsArray(responseObject[XmlElementNames_1.XmlElementNames.People], XmlElementNames_1.XmlElementNames.Persona);
        for (var _i = 0, responses_1 = responses; _i < responses_1.length; _i++) {
            this.people.push(responses_1[_i]);
        }
    };

    return FindPeopleResponse;
}(ServiceResponse_1.ServiceResponse));
exports.FindPeopleResponse = FindPeopleResponse;
