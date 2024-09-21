export interface Event {
    subject: string,
    start: TimeZoneDate,
    end: TimeZoneDate,
    recurrence?: recurrenceEvent,
    isOnlineMeeting?: boolean,
    allowNewTimeProposals?: boolean,
    onlineMeetingProvider?: string,
    attendees?:attendees[],
    location:{
        displayName :string | ""
    },
    extensions: [
        {
            "@odata.type": "microsoft.graph.openTypeExtension",
            extensionName: "Com.Contoso.Referral",
            appointmentType: string,
            clinicianName: string
            client:string,
            note:string
        }
    ]
}

export interface attendees {
    emailAddress: {
        address: string,
        name: string
    },
    type: string

}

export interface recurrenceEvent {
    pattern: {
        type: string,
        interval: number,
        daysOfWeek: string[],
    }
    range: {
        type: string,
        startDate: string,
        endDate: string
    }
}

export interface TimeZoneDate {
    dateTime: string,
    timeZone: string
}