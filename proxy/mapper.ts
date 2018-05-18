import { EnumValues } from "enum-values/src/enumValues";
import { Appointment, AppointmentSchema, AppointmentType, AttendeeCollection, BodyType, DateTime, DateTimeKind, DayOfTheWeek, DayOfTheWeekIndex, ExtendedPropertyDefinition, IOutParam, LegacyFreeBusyStatus, MeetingResponseType, MessageBody, PropertyDefinition, PropertyDefinitionBase, Recurrence, StringList, TimeZoneInfo } from "ews-javascript-api";
import * as moment from 'moment-timezone';
import * as xmlEscape from 'xml-escape';
import { EventAvailability, OfficeApiEvent, RecurrencePatternType, RecurrenceRangeType } from "../model/office";

//import { raw } from 'body-parser';

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
        appointment.Location = rawEvent.location ? xmlEscape(rawEvent.location.displayName) : '';
    }

    if (rawEvent.subject) {
        appointment.Subject = xmlEscape(rawEvent.subject);
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
        let startTimezone = TimeZoneInfo.FindSystemTimeZoneById(rawEvent.start.timeZone);
        let mDate = moment.tz(rawEvent.start.dateTime as string, startTimezone.IanaId);
        appointment.Start = new DateTime(mDate);
        appointment.Start.kind = DateTimeKind.Unspecified;
        appointment.StartTimeZone = startTimezone;
    }

    if (rawEvent.end) {
        let endTimezone = TimeZoneInfo.FindSystemTimeZoneById(rawEvent.end.timeZone);
        let mDate = moment.tz(rawEvent.end.dateTime as string, endTimezone.IanaId);
        appointment.End = new DateTime(mDate);
        appointment.End.kind = DateTimeKind.Unspecified;
        appointment.EndTimeZone = endTimezone;
    }

    if (rawEvent.recurrence) {
        //Set Recurrence is not typed correctly....
        if (rawEvent.recurrence.pattern && rawEvent.recurrence.range) {
            if (rawEvent.recurrence.pattern.type === RecurrencePatternType.Daily) {
                //@ts-ignore
                appointment.Recurrence = new Recurrence.DailyPattern(new DateTime(rawEvent.recurrence.range.startDate), rawEvent.recurrence.pattern.interval);
            } else if (rawEvent.recurrence.pattern.type === RecurrencePatternType.Weekly) {
                let daysOfWeekEnum = mapDaysOfTheWeek(rawEvent.recurrence.pattern.daysOfWeek);
                //@ts-ignore
                appointment.Recurrence = new Recurrence.WeeklyPattern(new DateTime(rawEvent.recurrence.range.startDate), rawEvent.recurrence.pattern.interval, daysOfWeekEnum);
            } else if (rawEvent.recurrence.pattern.type === RecurrencePatternType.RelativeMonthly) {
                //@ts-ignore
                appointment.Recurrence = new Recurrence.RelativeMonthlyPattern(new DateTime(rawEvent.recurrence.range.startDate), rawEvent.recurrence.pattern.interval, DayOfTheWeek[rawEvent.recurrence.pattern.daysOfWeek[0]], DayOfTheWeekIndex[rawEvent.recurrence.pattern.index]);
            } else if (rawEvent.recurrence.pattern.type === RecurrencePatternType.AbsoluteMonthly) {
                //@ts-ignore
                appointment.Recurrence = new Recurrence.MonthlyPattern(new DateTime(rawEvent.recurrence.startDate), rawEvent.recurrence.pattern.interval, rawEvent.recurrence.pattern.dayOfMonth);
            } else if (rawEvent.recurrence.pattern.type === RecurrencePatternType.RelativeYearly) {
                //@ts-ignore
                appointment.Recurrence = new Recurrence.RelativeYearlyPattern(new DateTime(rawEvent.recurrence.range.startDate), rawEvent.recurrence.pattern.month, DayOfTheWeek[rawEvent.recurrence.pattern.daysOfWeek[0]], DayOfTheWeekIndex[rawEvent.recurrence.pattern.index]);
            } else if (rawEvent.recurrence.pattern.type === RecurrencePatternType.AbsoluteYearly) {
                //@ts-ignore
                appointment.Recurrence = new Recurrence.YearlyPattern(new DateTime(rawEvent.recurrence.range.startDate), rawEvent.recurrence.pattern.month, rawEvent.recurrence.pattern.dayOfMonth);
            }
        }
        //Set Recurrence Range
        if (rawEvent.recurrence.range.type === RecurrenceRangeType.Numbered) {
            appointment.Recurrence.numberOfOccurrences = rawEvent.recurrence.range.numberOfOccurrences;
        } else if (rawEvent.recurrence.range.type === RecurrenceRangeType.EndDate) {
            appointment.Recurrence.endDate = new DateTime(rawEvent.recurrence.range.endDate);
        }
    }
}

export function mapDaysOfTheWeek(daysOfWeek): DayOfTheWeek[] {
    let daysOfWeekEnum = []
    daysOfWeek.map((day) => {
        daysOfWeekEnum.push(DayOfTheWeek[day]);
    });
    return daysOfWeekEnum;
}

export function mapAppointmentToApiEvent(item: Appointment, additionalProps?: PropertyDefinitionBase[]): OfficeApiEvent {
    if (!item)
        return null;

    let result: OfficeApiEvent = null;

    //@ts-ignore
    if (item.Sensitivity !== "Normal") {
        result = {
            id: item.Id.UniqueId,
            start: {
                dateTime: getOfficeDateTime(item.Start, item.StartTimeZone, item.IsAllDayEvent),
                timeZone: 'UTC'
            },
            end: {
                dateTime: getOfficeDateTime(item.End, item.StartTimeZone, item.IsAllDayEvent),
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
            subject: item.Subject,
            start: {
                dateTime: getOfficeDateTime(item.Start, item.StartTimeZone, item.IsAllDayEvent),
                timeZone: 'UTC'
            },
            end: {
                dateTime: getOfficeDateTime(item.End, item.StartTimeZone, item.IsAllDayEvent),
                timeZone: 'UTC'
            },
            location: { displayName: item.Location },
            isAllDay: item.IsAllDayEvent,
            //@ts-ignore
            showAs: getFreeBusyStatusNewName(LegacyFreeBusyStatus[item.LegacyFreeBusyStatus]),
            type: getAppointmentType(<any>item.AppointmentType),
            seriesMasterId: (isSeriesItem(item)) ? "masterFor:" + item.Id.UniqueId : undefined,
            sensitivity: <any>item.Sensitivity,
            isMeeting: item.IsMeeting
        };

        if (item.GetLoadedPropertyDefinitions().find((p: PropertyDefinition) => p.Name === AppointmentSchema.WebClientReadFormQueryString.Name)) {
            let webLink = '';
            //According to https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.item.webclientreadformquerystring(v=exchg.80).aspx
            if (item.WebClientReadFormQueryString.indexOf('http') === 0) {
                webLink = item.WebClientReadFormQueryString;
            } else {
                webLink = `${item.Service.Url.Scheme}://${item.Service.Url.Host}/owa/${item.WebClientReadFormQueryString}`;
            }
            result.webLink = webLink;
        }

        if (hasProperty(item, AppointmentSchema.ParentFolderId)) {
            result.calendarId = item.ParentFolderId.UniqueId;
        }

        if (hasProperty(item, AppointmentSchema.IsReminderSet)) {
            result.isReminderOn = item.IsReminderSet;
        }

        if (hasProperty(item, AppointmentSchema.ReminderMinutesBeforeStart)) {
            result.reminderMinutesBeforeStart = item.ReminderMinutesBeforeStart;
        }

        if (hasProperty(item, AppointmentSchema.Organizer)) {
            result.organizer = ({
                emailAddress: {
                    name: item.Organizer.Name,
                    address: item.Organizer.Address
                }
            });
        }

        if (hasProperty(item, AppointmentSchema.Body)) {
            result.body = ({
                contentType: EnumValues.getNameFromValue(BodyType, item.Body.BodyType),
                content: item.Body.Text
            });
        }

        if (item.RequiredAttendees.Count >= 1 || item.OptionalAttendees.Count >= 1 || item.Resources.Count >= 1) {
            let attendees = mapAttendees(item.RequiredAttendees, "Required");
            let withOptional = attendees.concat(mapAttendees(item.OptionalAttendees, "Optional"));
            let all = withOptional.concat(mapAttendees(item.Resources, "Resource"));
            result.attendees = all;
        }

        if (hasProperty(item, AppointmentSchema.Body) && item.Categories.Count > 0) {
            result.categories = item.Categories.GetEnumerator();
        }
    }

    if (additionalProps && additionalProps.length > 0) {
        additionalProps.forEach(prop => {
            if (prop instanceof ExtendedPropertyDefinition) {
                let outParam: IOutParam<any> = {
                    outValue: null
                };

                item.TryGetProperty<any>(prop, outParam);
                result[prop.Name[0].toLowerCase() + prop.Name.substring(1)] = outParam.outValue;
            }
        });
    }

    return result;
}

export function getOfficeDateTime(date: DateTime, timezone: TimeZoneInfo, isAllDay: boolean): string {
    let result: string;
    // Some background.. This is mimicking the Graph API behavior for all day events
    // => The just return an UTC date @00:00:00, which is technically not the same
    // that is saved on Exchange, but much easier to consume. So we do the same ->
    // Interpret the dateTime of the all-day event in the respective timezone, which
    // should lead to a date @00:00:00. Then, format this as a local date time and
    // tell the consumer it's in UTC
    if (isAllDay) {
        result = moment(date.ToISOString()).tz(timezone.IanaId).format("YYYY-MM-DDTHH:mm:ss");
    } else {
        // To-Do: Actually supply the correct timezone.. We need this to keep the events
        // in the correct timezone for updates
        result = moment.utc(date.ToISOString()).format("YYYY-MM-DDTHH:mm:ss");
    }

    return result;
}

export function hasProperty(item: Appointment, property: PropertyDefinition) {
    if (item.GetLoadedPropertyDefinitions().find((p: PropertyDefinition) => p.Name === property.Name)) {
        return true;
    } else {
        return false;
    }
}

export function isNullOrEmpty(s: string): boolean {
    return s == null || s.length === 0;
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

export function isSeriesItem(appointment: Appointment) {
    if (typeof (appointment.AppointmentType) === 'string') {
        return appointment.AppointmentType === 'Occurrence' || appointment.AppointmentType === 'Exception';
    } else {
        return appointment.AppointmentType === AppointmentType.Exception || appointment.AppointmentType === AppointmentType.Occurrence;
    }
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