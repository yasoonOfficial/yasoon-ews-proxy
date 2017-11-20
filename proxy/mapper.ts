import * as moment from 'moment-timezone';
import * as xmlEscape from 'xml-escape';

import { OfficeApiEvent, EventAvailability } from "../model/office";
import { Appointment, BodyType, MessageBody, StringList, DateTime, DateTimeKind, AttendeeCollection, MeetingResponseType, LegacyFreeBusyStatus } from "ews-javascript-api";
import { EnumValues } from "enum-values/src/enumValues";


export function copyApiEventToAppointment(rawEvent: OfficeApiEvent, appointment: Appointment) {
    //Let the mapping begin!
    if (rawEvent.attendees) {
        rawEvent.attendees.forEach(a => {
            if (!a.type || a.type === 'Required') {
                appointment.RequiredAttendees.Add(a.emailAddress.address);
            } else if (a.type === 'Optional') {
                appointment.OptionalAttendees.Add(a.emailAddress.address);
            } else if (a.type === 'Resource') {
                appointment.Resources.Add(a.emailAddress.address);
            }
        });
    }

    if (rawEvent.body) {
        let bodyType = BodyType.Text;
        if (rawEvent.body.contentType.toLowerCase() === 'html')
            bodyType = BodyType.HTML;

        //#shitty api
        if (bodyType === BodyType.HTML)
            rawEvent.body.content = xmlEscape(rawEvent.body.content);

        appointment.Body = new MessageBody(bodyType, rawEvent.body.content);
    }

    if (rawEvent.showAs) {
        //@ts-ignore
        appointment.LegacyFreeBusyStatus = EnumValues.getNameFromValue(LegacyFreeBusyStatus, getLegacyFreeBusyStatusFromName(rawEvent.showAs));
    }

    if (rawEvent.categories) {
        appointment.Categories = new StringList(rawEvent.categories || []);
    }

    if (rawEvent.location) {
        appointment.Location = rawEvent.location ? rawEvent.location.displayName : '';
    }

    if (rawEvent.subject) {
        appointment.Subject = rawEvent.subject;
    }

    if (rawEvent.isAllDay !== undefined) {
        appointment.IsAllDayEvent = rawEvent.isAllDay;
    }

    if (rawEvent.isReminderOn === true) {
        appointment.IsReminderSet = true;
        appointment.ReminderMinutesBeforeStart = rawEvent.reminderMinutesBeforeStart;
    } else if (rawEvent.isReminderOn === false) {
        appointment.IsReminderSet = false;
    }

    if (rawEvent.start) {
        let mDate = moment.tz(rawEvent.start.dateTime as string, rawEvent.start.timeZone);
        appointment.Start = new DateTime(mDate);
        appointment.Start.kind = DateTimeKind.Unspecified;
        //appointment.StartTimeZone = TimeZoneInfo.FindSystemTimeZoneById(rawEvent.start.timeZone);
    }

    if (rawEvent.end) {
        let mDate = moment.tz(rawEvent.end.dateTime as string, rawEvent.end.timeZone);
        appointment.End = new DateTime(mDate);
        appointment.End.kind = DateTimeKind.Unspecified;
        //appointment.EndTimeZone = TimeZoneInfo.FindSystemTimeZoneById(rawEvent.end.timeZone);
    }
}

export function mapAppointmentToApiEvent(item: Appointment): OfficeApiEvent {
    if (!item)
        return null;

    let attendees = mapAttendees(item.RequiredAttendees, "Required");
    let withOptional = attendees.concat(mapAttendees(item.OptionalAttendees, "Optional"));
    let all = withOptional.concat(mapAttendees(item.Resources, "Resource"));

    let result: OfficeApiEvent = null;

    //@ts-ignore
    if (item.Sensitivity !== "Normal") {
        result = {
            id: item.Id.UniqueId,
            start: {
                dateTime: item.Start.ToISOString(),
                timeZone: 'UTC'
            },
            end: {
                dateTime: item.End.ToISOString(),
                timeZone: 'UTC'
            },
            type: 'singleInstance',
            isAllDay: item.IsAllDayEvent,
            sensitivity: <any>item.Sensitivity,
            subject: <any>item.Sensitivity,
            //@ts-ignore
            showAs: getFreeBusyStatusNewName(LegacyFreeBusyStatus[item.LegacyFreeBusyStatus])
        };
    } else {
        result = {
            id: item.Id.UniqueId,
            calendarId: (item.ParentFolderId ? item.ParentFolderId.UniqueId : ''),
            subject: item.Subject,
            start: {
                dateTime: item.Start.ToISOString(),
                timeZone: 'UTC'
            },
            end: {
                dateTime: item.End.ToISOString(),
                timeZone: 'UTC'
            },
            location: { displayName: item.Location },
            isAllDay: item.IsAllDayEvent,
            //@ts-ignore
            showAs: getFreeBusyStatusNewName(LegacyFreeBusyStatus[item.LegacyFreeBusyStatus]),
            categories: (item.Categories ? item.Categories.GetEnumerator() : []),
            organizer: (item.Organizer ? ({
                emailAddress: {
                    name: item.Organizer.Name,
                    address: item.Organizer.Address
                }
            }) : null),
            type: getAppointmentType(<any>item.AppointmentType),
            isReminderOn: item.IsReminderSet,
            reminderMinutesBeforeStart: (item.IsReminderSet) ? item.ReminderMinutesBeforeStart : undefined,
            attendees: all,
            sensitivity: <any>item.Sensitivity,
            body: (item.Body ? ({
                contentType: EnumValues.getNameFromValue(BodyType, item.Body.BodyType),
                content: item.Body.Text
            }) : null)
        };
    }

    return result;
}


export function mapAttendees(attendees: AttendeeCollection, type: string) {
    return attendees.GetEnumerator().map(a => ({
        type: type,
        status: {
            response: getResponseStatusName(a.ResponseType),
            time: a.LastResponseTime ? a.LastResponseTime.ToISOString() : null
        },
        emailAddress: {
            name: a.Name,
            address: a.Address
        }
    }));
}


export function getAppointmentType(type: string) {
    switch (type) {
        case "Single":
            return "singleInstance";
        case "Occurrence":
            return "occurrence";
        case "Exception":
            return "exception";
        case "RecurringMaster":
            return "seriesMaster";
    }
}

export function getResponseStatusName(type: MeetingResponseType) {
    switch (type) {
        case MeetingResponseType.Accept:
            return "Accepted";
        case MeetingResponseType.Decline:
            return "Declined";
        case MeetingResponseType.NoResponseReceived:
            return "NotResponded";
        case MeetingResponseType.Organizer:
            return "Organizer";
        case MeetingResponseType.Tentative:
            return "TentativelyAccepted";
        case MeetingResponseType.Unknown:
            return "None";
    }
}

export function getFreeBusyStatusLabel(status: LegacyFreeBusyStatus): string {
    switch (status) {
        case LegacyFreeBusyStatus.Busy:
            return 'Busy';
        case LegacyFreeBusyStatus.Free:
            return 'Free';
        case LegacyFreeBusyStatus.NoData:
            return 'Unknown';
        case LegacyFreeBusyStatus.OOF:
            return 'Out of Office';
        case LegacyFreeBusyStatus.Tentative:
            return 'Tentative';
        case LegacyFreeBusyStatus.WorkingElsewhere:
            return 'Working Elsewhere';
    }
}


export function getLegacyFreeBusyStatusFromName(name: string): LegacyFreeBusyStatus {
    switch (name) {
        case 'busy':
            return LegacyFreeBusyStatus.Busy;
        case 'free':
            return LegacyFreeBusyStatus.Free;
        case 'unknown':
            return LegacyFreeBusyStatus.NoData;
        case 'oof':
            return LegacyFreeBusyStatus.OOF;
        case 'tentative':
            return LegacyFreeBusyStatus.Tentative;
        case 'workingElsewhere':
            return LegacyFreeBusyStatus.WorkingElsewhere;
    }
}

export function getFreeBusyStatusNewName(status: LegacyFreeBusyStatus): EventAvailability {
    if (status === null)
        return null;

    switch (status) {
        case LegacyFreeBusyStatus.Busy:
            return <EventAvailability>'busy';
        case LegacyFreeBusyStatus.Free:
            return <EventAvailability>'free';
        case LegacyFreeBusyStatus.NoData:
            return <EventAvailability>'unknown';
        case LegacyFreeBusyStatus.OOF:
            return <EventAvailability>'oof';
        case LegacyFreeBusyStatus.Tentative:
            return <EventAvailability>'tentative';
        case LegacyFreeBusyStatus.WorkingElsewhere:
            return <EventAvailability>'workingElsewhere';
    }
}