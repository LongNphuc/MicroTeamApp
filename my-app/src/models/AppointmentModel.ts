export interface AppointmentModel{
    id: string,
    subject: string,
    start: TimeZoneDate,
    end: TimeZoneDate,
    isCancelled: string,
    organizer: {
        emailAddress: {
            name: string,
            address: string
        }
    },
    onlineMeeting: string | undefined,
    location: {
        displayName: string
    },
    extensions: [
        {
            appointmentType: string,
            clinicianName: string,
            client: string,
            note:string
        }
    ]
}

export interface TimeZoneDate{
    dateTime : Date,
    timeZone : string
}

export interface UTCDateRange{
    start: string,
    end:string
}