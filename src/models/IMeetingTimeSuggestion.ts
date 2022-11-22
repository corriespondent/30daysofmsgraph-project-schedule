export interface IMeetingTimeSuggestion {
    attendeeAvailability: Array<any>,
    confidence: number,
    locations: Array<any>,
    meetingTimeSlot: {
        end: { 
            dateTime: string,
            timeZone: string
        },
        start: {
            dateTime: string,
            timeZone: string
        }
    },
    organizerAvailability: string
}