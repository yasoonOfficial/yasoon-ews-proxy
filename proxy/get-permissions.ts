import { AttendeeInfo, AvailabilityData, BasePropertySet, DateTime, ExchangeService, ExchangeVersion, Folder, FolderId, FolderSchema, GetUserAvailabilityResults, Mailbox, PropertySet, ServiceResult, TimeWindow, TimeZoneInfo, Uri, WellKnownFolderName } from "ews-javascript-api";
import { Environment } from "../model/proxy";
import { applyCredentials, getAccessArrayFromEffectiveRights } from "../proxy/helper";


export interface GetPermissionsParams {
    email: string;
    calendarId: string;
}

export class GetPermissionsRequest {

    async execute(env: Environment, params: GetPermissionsParams) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let targetFolderId = null;
        if (params.calendarId === 'main') {
            targetFolderId = new FolderId(WellKnownFolderName.Calendar, new Mailbox(params.email));
        } else {
            targetFolderId = new FolderId();
            targetFolderId.UniqueId = params.calendarId;
        }

        let targetCalendar: Folder;

        //First check if we have read access to this calendar
        try {
            targetCalendar = await Folder.Bind(service, targetFolderId, new PropertySet(BasePropertySet.IdOnly, FolderSchema.EffectiveRights));
        }
        catch (e) {
            console.log(e.message);
        }

        if (targetCalendar && targetCalendar.Id && targetCalendar.Id.UniqueId) {
            //We got at least full read access, check the rest
            if (targetCalendar.EffectiveRights) {
                let access = getAccessArrayFromEffectiveRights(targetCalendar.EffectiveRights);
                //If there is no access as of now, we only have free-busy access! Not sure how we were able to get the
                // folder ID that way, but well... Happens for timur :D
                if (access.length > 0) {
                    return {
                        id: targetCalendar.Id.UniqueId,
                        access: access
                    };
                } //Else -> Fall back to free-busy, see below
            } //Else -> Fall back to free-busy, see below
        }

        //We don't have full read access, check if we can get free busy data
        let attendee = new AttendeeInfo(params.email);
        let availability: GetUserAvailabilityResults;

        try {
            availability = await service.GetUserAvailability([attendee], new TimeWindow(DateTime.Now, DateTime.Now.AddDays(1)), AvailabilityData.FreeBusy);
        } catch (e) {
            console.log(e.message);
            console.log(e.stack);
        }

        if (availability && availability.AttendeesAvailability && availability.AttendeesAvailability.Responses &&
            availability.AttendeesAvailability.Responses[0].Result !== ServiceResult.Error) {
            return {
                access: ['freebusy']
            };
        }

        return {
            access: ['none']
        };
    }
}