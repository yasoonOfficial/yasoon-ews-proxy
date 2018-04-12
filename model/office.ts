import { DayOfTheWeek, DayOfTheWeekIndex } from "ews-javascript-api";

export interface OfficeApiEvent {
    id?: string;
    calendarId?: string;
    subject?: string;
    start?: {
        dateTime: string;
        timeZone: string;
    };
    isReminderOn?: boolean;
    reminderMinutesBeforeStart?: number;
    sensitivity?: string;
    end?: {
        dateTime: string;
        timeZone: string;
    };
    isAllDay?: boolean;
    color?: string;
    categories?: string[];
    type?: string;
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
    recurrence?: {
        pattern: RecurrencePattern;
        range: RecurrenceRange;
    };
    seriesMasterId?: string;
    isMeeting?: boolean;
}


export interface SingleValueExtendedProperty {
    id: string;
    value: string;
}

export interface XmlCategoriesResult {
    categories: XmlCategories;
}

export interface XmlCategories {
    category: XmlCategory[];
}

export interface XmlCategory {
    $: {
        color: string;
        guid: string;
        keyboardShortcut: string;
        lastSessionUsed: string;
        lastTimeUsed: string;
        lastTimeUsedCalendar: string;
        lastTimeUsedContacts: string;
        lastTimeUsedJournal: string;
        lastTimeUsedMail: string;
        lastTimeUsedNotes: string;
        lastTimeUsedTasks: string;
        name: string;
        renameOnFirstUse: string;
        usageCount: string;
    }
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

export interface PatternedRecurrence {
    pattern: RecurrencePattern
    range: RecurrenceRange
}

export interface RecurrencePattern {
    dayOfMonth?: number;
    daysOfWeek?: DayOfTheWeek[];
    firstDayOfWeek?: string;
    index?: DayOfTheWeekIndex;
    interval?: number;
    month?: number;
    type?: RecurrencePatternType;
}

export interface RecurrenceRange {
    endDate?: string;    //"String (timestamp)"
    numberOfOccurrences?: number;
    recurrenceTimeZone?: string;
    startDate?: string;  //"String (timestamp)"
    type?: RecurrenceRangeType;
}

export enum RecurrencePatternType {
    Daily = 'Daily', Weekly = 'Weekly', AbsoluteMonthly = 'AbsoluteMonthly',
    RelativeMonthly = 'RelativeMonthly', AbsoluteYearly = 'AbsoluteYearly',
    RelativeYearly = 'RelativeYearly'
}

export enum RecurrenceRangeType {
    EndDate = 'EndDate', NoEnd = 'NoEnd', Numbered = 'Numbered'
}