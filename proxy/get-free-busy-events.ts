import * as moment from 'moment-timezone';

import { Environment } from "../model/proxy";
import {
    ExchangeService, Uri, ExchangeVersion, TimeZoneInfo,
    AttendeeInfo, DateTime, AvailabilityOptions, FreeBusyViewType,
    TimeWindow, AvailabilityData, ServiceResult
} from "ews-javascript-api";
import { applyCredentials } from "../proxy/helper";
import { getFreeBusyStatusNewName, getFreeBusyStatusLabel } from '../proxy/mapper';
import { OfficeApiEvent } from '../model/office';

export interface GetFreeBusyEventsParams {
    email: string;
    startDate: string;
    endDate: string;
}

export class GetFreeBusyEventsRequest {

    async execute(env: Environment, params: GetFreeBusyEventsParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        //We don't have full read access, check if we can get free busy data
        let attendee = new AttendeeInfo(params.email);
        let startDate = new DateTime(moment(params.startDate as string));
        let endDate = new DateTime(moment(params.endDate as string));

        //Request as much information as possible, subject and location may be set!
        let options = new AvailabilityOptions();
        options.RequestedFreeBusyView = FreeBusyViewType.DetailedMerged;

        let availability = await service.GetUserAvailability([attendee], new TimeWindow(startDate, endDate), AvailabilityData.FreeBusy, options);
        if (availability.AttendeesAvailability.Responses[0].Result === ServiceResult.Error) {
            throw new Error(availability.AttendeesAvailability.Responses[0].ErrorMessage);
        }

        let calendarEvents = availability.AttendeesAvailability.Responses[0].CalendarEvents;

        return calendarEvents.map(c => {
            let id = 'freeBusy' + c.StartTime.ToISOString();
            let location = '';
            let subject = getFreeBusyStatusLabel(c.FreeBusyStatus);

            if (c.Details) {
                id = c.Details.StoreId || id;
                location = c.Details.Location || location;
                subject = c.Details.Subject || subject;
            }

            return <OfficeApiEvent>{
                id: id,
                calendarId: 'main',
                start: {
                    dateTime: c.StartTime.ToISOString(),
                    timeZone: 'UTC'
                },
                end: {
                    dateTime: c.EndTime.ToISOString(),
                    timeZone: 'UTC'
                },
                subject: subject,
                location: { displayName: location },
                isAllDay: (c.EndTime.Subtract(c.StartTime).TotalHours >= 24),
                showAs: getFreeBusyStatusNewName(c.FreeBusyStatus)
            };
        });
    }
}