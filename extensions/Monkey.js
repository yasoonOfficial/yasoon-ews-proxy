"use strict";
Object.defineProperty(exports, "__esModule", { value: true });

var TimeZonePropertyDefinition_1 = require("ews-javascript-api/js/PropertyDefinitions/TimeZonePropertyDefinition");
var TimeZoneDefinition_1 = require("ews-javascript-api/js/ComplexProperties/TimeZones/TimeZoneDefinition");
var CreateItemRequestBase = require("ews-javascript-api/js/Core/Requests/CreateItemRequestBase");

var Monkey = (function (_super) {
    function Monkey() {
    }

    Monkey.prototype.patch = function () {
        TimeZonePropertyDefinition_1.TimeZonePropertyDefinition.prototype.WritePropertyValueToXml = function (writer, propertyBag, isUpdateOperation) {
            var value = propertyBag._getItem(this);
            if (value != null) {
                // We emit time zone properties only if we have not emitted the time zone SOAP header
                // or if this time zone is different from that of the service through which the request
                // is being emitted.
                if (!writer.IsTimeZoneHeaderEmitted || value != writer.Service.TimeZone) {
                    var timeZoneDefinition = new TimeZoneDefinition_1.TimeZoneDefinition(value);
                    writer.WriteStartElement(timeZoneDefinition.Namespace, this.XmlElementName);
                    timeZoneDefinition.WriteAttributesToXml(writer);
                    timeZoneDefinition.WriteElementsToXml(writer);
                    writer.WriteEndElement();
                }
            }
        };

        CreateItemRequestBase.prototype.WriteAttributesToXml = function (writer) {
            _super.prototype.WriteAttributesToXml.call(this, writer);
            if (this.MessageDisposition !== null) {
                writer.WriteAttributeValue(XmlAttributeNames_1.XmlAttributeNames.MessageDisposition, MessageDisposition_1.MessageDisposition[this.MessageDisposition]);
            }
            if (this.SendInvitationsMode !== null) {
                writer.WriteAttributeValue(XmlAttributeNames_1.XmlAttributeNames.SendMeetingInvitations, SendInvitationsMode_1.SendInvitationsMode[this.SendInvitationsMode]);
            }
            if (this.Service.SkipSendingMeetingInviteToGroup !== null) {
                writer.WriteAttributeValue("SkipSendingMeetingInviteToGroup", this.Service.SkipSendingMeetingInviteToGroup);
            }
        };
    }

    return Monkey;
})();

exports.Monkey = Monkey;