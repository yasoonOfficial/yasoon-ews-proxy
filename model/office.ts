export interface OfficeApiEvent {
    id?: string;
    calendarId?: string;
    subject?: string;
    start?: {
        dateTime: string;
        timeZone: string;
    };

    end?: {
        dateTime: string;
        timeZone: string;
    };
    isAllDay?: boolean;
    color?: string;
    categories?: string[];
    location?: {
        displayName: string;
        address?: string;
    };
    attendees?: OfficeEventAttendee[];
    organizer?: OfficeEventOrganizer;
    body?: {
        contentType: string;
        content: string;
    },
    showAs?: EventAvailability;
    webLink?: string;
    senderName?: string;
    singleValueExtendedProperties?: SingleValueExtendedProperty[];
}


export interface SingleValueExtendedProperty {
    id: string;
    value: string;
}

export interface OfficeEventAttendee {
    type?: string;
    status?: {
        response?: string;
        time?: string;
    };

    emailAddress: {
        name: string;
        address: string; //Email Address
    };
}

export interface OfficeEventOrganizer {
    emailAddress: {
        name: string;
        address: string; //Email Address
    };
}

export interface OfficeSharedMailbox {
    smtpAddress: string;
    name: string;
}

export enum EventAvailability {
    Busy = "busy", Free = "free", NoData = "unknown", Tentative = "tentative", OutOfOffice = "oof", WorkingElsewhere = "workingElsewhere"
}