import * as React from 'react';
import styles from './Scheduler.module.scss';
import { ISchedulerProps } from './ISchedulerProps';
import { ISchedulerState } from './ISchedulerState';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IUser } from '../../../models/IUser';
import { PrimaryButton, Dropdown } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/DateTimePicker';

export default class Scheduler extends React.Component<ISchedulerProps, ISchedulerState> {
  constructor(props: ISchedulerProps, state: ISchedulerState){
    super(props);
    this.state = {
      users: []
    }
  }

  public render(): React.ReactElement<ISchedulerProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      context
    } = this.props;


    return (
      <section>
        <div>
          <h1>Find a time to meet</h1>
          <div className={styles.formElement}>
            <PeoplePicker
              context={this.props.context as any}
              titleText="Choose people"
              showtooltip={true}
              personSelectionLimit={100}
              required={true}
              onChange={this._updatePeoplePickerItems}
            />
          </div>
          <p>Set the time block to look for meetings, or leave these blank to find the next available time.</p>
          <div className={styles.formElement}>
            <DateTimePicker
              label="Start time block"
              dateConvention={DateConvention.DateTime}
              timeConvention={TimeConvention.Hours12}
              timeDisplayControlType={TimeDisplayControlType.Dropdown}
              showSeconds={false}
              showClearDate={true}
              onChange={this._updateStart}
            />
            <DateTimePicker
              label="End time block"
              dateConvention={DateConvention.DateTime}
              timeConvention={TimeConvention.Hours12}
              timeDisplayControlType={TimeDisplayControlType.Dropdown}
              showSeconds={false}
              showClearDate={true}
              onChange={this._updateEnd}
              disabled={!this.state.start}
              onGetErrorMessage={(value) => this._validateEnd(value)}
            />
          </div>
          <div className={styles.formElement}>
            <Dropdown
              placeholder="Select meeting length"
              label="How long is your meeting?"
              options={[
                { key: "PT15M", text: "15 minutes"},
                { key: "PT30M", text: "30 minutes"},
                { key: "PT1H", text: "1 hour"},
                { key: "PT1H30M", text: "1 hour 30 minutes"},
                { key: "PT2H", text: "2 hours"},
                { key: "PT2H30M", text: "2 hours 30 minutes"},
                { key: "PT3H", text: "3 hours"}
              ]}
              defaultSelectedKey={ !this.state.meetingDuration ? "PT30M" : this.state.meetingDuration }
              selectedKey={this.state.meetingDuration}
              onChange={this._updateDuration}
            />
          </div>
          <div className={styles.formElement}>
            <PrimaryButton
              text="Look for meeting times"
              title="Look for meeting times"
              onClick={this._getMeetingTimes}
            />
          </div>
        </div>
        { console.log(this.state.meetingTimeSuggestions, this.state.emptySuggestionsReason)}
        { this.state.meetingTimeSuggestions && this.state.meetingTimeSuggestions.length > 0 && (
          <div>
            <h2>Everyone is available to meet:</h2>
            <ul>
            { this.state.meetingTimeSuggestions.map( meetingTime => (
              <li>{this._formatDate(meetingTime.meetingTimeSlot.start.dateTime + 'Z', 'date')} {this._formatDate(meetingTime.meetingTimeSlot.start.dateTime + 'Z', 'time')} - {this._formatDate(meetingTime.meetingTimeSlot.end.dateTime + 'Z', 'time')}</li>
            ))}
            </ul>
          </div>
        )}
      </section>
    );
  }

  private _formatDate = (inputDate: string, format: string): string => {
    const date = new Date(inputDate);
    if( format =="date"){
      return date.toLocaleDateString();
    }
    else {
      return date.toLocaleTimeString();
    }
  }

  private _updatePeoplePickerItems = (items:IUser[]): void => {
    console.log('Items', items);
    this.setState({users: items});
  }

  private _updateStart = (start: any): void => {
    console.log('Start', start);
    this.setState({ start: start});
  }

  private _updateEnd = (end: any): void => {
    console.log('End', end);
    this.setState({ end: end});
  }

  private _updateDuration = ( event: any, option: any): void => {
    console.log("Duration", option);
    this.setState({ meetingDuration: option.key });
  }

  private _validateEnd = (value: Date): string => {
    const start = new Date(this.state.start);
    if( start > value){
      return "Please check your dates - start is after end";
    } else {
      return '';
    }
  }

  private _getMeetingTimes = (): void => {
    console.log("Get meeting times!");

    // format attendee list
    let attendeeList: any = [];
    if( this.state.users.length > 0){
      this.state.users.forEach( (user: IUser) => {
        attendeeList.push({
          "emailAddress": {
            "address": user.secondaryText,
            "name": user.text
          },
          "type": "Required"
        })
      });

      console.log("attendeeList", attendeeList);

      // check for start/end times

      let timeSlots: any = [];
      if( this.state.start && this.state.end ){
        // timeSlots.push()
        const startDateTime = new Date(this.state.start);
        const endDateTime = new Date(this.state.end);
        const myTimeZone = new Date().toString().split('(')[1].split(')')[0];
        timeSlots.push({
          "start": {
            "dateTime": startDateTime.toISOString(),
            "timeZone": myTimeZone
          },
          "end": {
            "dateTime": endDateTime.toISOString(),
            "timeZone": myTimeZone
          }
        });
        console.log("timeSlots", timeSlots);
      }
    
      const requestBody: any = {
        "attendees": attendeeList,
        "timeConstraint": {
            "timeslots": timeSlots
        },
        "locationConstraint": {
            "isRequired": "false",
            "suggestLocation": "false",
            "locations": []
        },
        "meetingDuration": this.state.meetingDuration,
        "returnSuggestionReasons": "true",
        "minimumAttendeePercentage": "100",
        "maxCandidates": "10"
      }

      this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
        .api("me/findMeetingTimes")
        .post(requestBody, (err, res) => {
          if(err){
            console.error(err);
            return;
          }

          console.log("response", res);
          this.setState({ 
            meetingTimeSuggestions: res.meetingTimeSuggestions,
            emptySuggestionsReason: res.emptySuggestionsReason 
          });
        });
      });
    }
  }

}
