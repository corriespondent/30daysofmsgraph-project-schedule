import { IUser } from '../../../models/IUser';
import { IMeetingTimeSuggestion } from '../../../models/IMeetingTimeSuggestion';

export interface ISchedulerState {
    users: Array<IUser>,
    start?: string,
    end?: string,
    emptySuggestionsReason?: string,
    meetingTimeSuggestions?: Array<IMeetingTimeSuggestion>,
    meetingDuration?: string
}