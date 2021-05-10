// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as React from 'react';
import styles from './MyCalendar.module.scss';
import * as strings from 'MyCalendarWebPartStrings';
import { IMeeting, IMeetings } from './IMeeting';
import { IMyCalendarState } from './IMyCalendarState';
import { IMyCalendarProps } from './IMyCalendarProps';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { Link } from '@fluentui/react/lib/components/Link';
import { List } from '@fluentui/react/lib/components/List';
import { PrimaryButton, IIconProps } from '@fluentui/react';
import { initializeIcons,  } from '@uifabric/icons';
import { mergeStyles, mergeStyleSets } from '@fluentui/react/lib/Styling';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from '@fluentui/react/lib/Stack';
import { Text, ITextProps } from '@fluentui/react/lib/Text';
import { Separator } from '@fluentui/react/lib/Separator';
import { Event } from '@microsoft/microsoft-graph-types';
import { GraphRequest } from "@microsoft/sp-http";
import { unstable_renderSubtreeIntoContainer } from 'react-dom';

export default class MyCalendar extends React.Component<IMyCalendarProps, IMyCalendarState> {
  private _interval: number;

  private _messageIconClass: string = mergeStyles({
    fontSize: 16,
    height: 16,
    width: 15,
    margin: '0',
  });

  constructor(props: IMyCalendarProps) {
    super(props);

    initializeIcons();

    console.log('contructor');

    this.state = {
      error: undefined,
      meetings: [],
      loading: true,
      renderedDateTime: new Date(),
      isOpen: false,
      activeEvent: {} as microsoftgraph.Event
    };
  }

  /**
   * Get timezone for logged in user
   */
  private _getTimeZone(): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      this.props.graphClient
        // get the mailbox settings
        .api(`me/mailboxSettings`)
        .version("v1.0")
        .get((err: any, res: microsoftgraph.MailboxSettings): void => {
          resolve(res.timeZone);
        });
    });
  }

  /**
     * Load recent messages for the current user
     */
  // private _loadMeetings(): void {
  //   if (!this.props.graphClient) {
  //     return;
  //   }

  //   // update state to indicate loading and remove any previously loaded
  //   // meetings
  //   this.setState({
  //     error: null,
  //     loading: true,
  //     meetings: []
  //   });

  //   const date: Date = new Date();
  //   const now: string = date.toISOString();
  //   // set the date to midnight today to load all upcoming meetings for today
  //   date.setHours(23);
  //   date.setMinutes(59);
  //   date.setSeconds(0);
  //   date.setDate(date.getDate() + (this.props.daysInAdvance || 0));
  //   const midnight: string = date.toISOString();

  //   this._getTimeZone().then(timeZone => {
  //     this.props.graphClient
  //       // get all upcoming meetings for the rest of the day today
  //       .api(`me/calendar/calendarView?startDateTime=${now}&endDateTime=${midnight}`)
  //       .version("v1.0")
  //       .select('subject,start,end,showAs,webLink,location,isAllDay')
  //       .top(this.props.numMeetings > 0 ? this.props.numMeetings : 100)
  //       .header("Prefer", "outlook.timezone=" + '"' + timeZone + '"')
  //       // sort ascending by start time
  //       .orderby("start/dateTime")
  //       .get((err: any, res: IMeetings): void => {
  //         if (err) {
  //           // Something failed calling the MS Graph
  //           this.setState({
  //             error: err.message ? err.message : strings.Error,
  //             loading: false
  //           });
  //           return;
  //         }

  //         // Check if a response was retrieved
  //         if (res && res.value && res.value.length > 0) {
  //           this.setState({
  //             meetings: res.value,
  //             loading: false
  //           });
  //         }
  //         else {
  //           // No meetings found
  //           this.setState({
  //             loading: false
  //           });
  //         }
  //       });
  //   });
  // }

  private _loadCalendar(): void {
    if (!this.props.graphClient) {
      return;
    }

    // update state to indicate loading and remove any previously loaded
    // meetings
    this.setState({
      error: null,
      loading: true,
      meetings: []
    });    

    const request: GraphRequest = this.props.graphClient
      .api(`me/calendar/calendarView?startDateTime=${new Date().toISOString()}&endDateTime=${this._getFilterEndDateString()}`)
      .version("v1.0")
      .select('subject,start,end,showAs,webLink,location,isAllDay')
      .top(this.props.numMeetings > 0 ? this.props.numMeetings : 100)
      .orderby("start/dateTime");

    console.log(request.buildFullUrl());

    request
      .get()
      .then((result: IMeetings) => {
        const meetings: IMeeting[] = (result && result.value) ? result.value : [];
        this.setState({
          meetings: meetings,
          loading: false
        });
      })
      .catch((err) => {
        this.setState({
          error: err.message ? err.message : strings.Error,
          loading: false
        });
      });      
  }

    

  /**
     * Render meeting item
     */
  private _onRenderCell = (item: IMeeting, index: number | undefined): JSX.Element => {
    const startDate: Date = new Date(item.start.dateTime + 'Z');
    const hour: number = startDate.getHours();
    const adjHour: number = (hour > 12) ? hour - 12 : hour;
    const amPm: string = ((hour >= 12) && (hour != 24)) ? "PM" : "AM";
    const minutes: number = startDate.getMinutes();
    
    const itemElement: JSX.Element = 
      <div className={`${styles.meetingWrapper} ${item.showAs}`}>
        <Link className={styles.meeting} onClick={() => this._showEventDetails(item.id)} target='_blank' >
          <Text nowrap block className={styles.start}>{`${adjHour}:${minutes < 10 ? '0' + minutes : minutes} ${amPm}`}</Text>
          <Text nowrap block className={styles.subject}>{item.subject}</Text>
          <Text nowrap block className={styles.duration}>{this._getDuration(item)}</Text>
          <Text nowrap block className={styles.location}>{item.location.displayName}</Text>                      
        </Link>
      </div>;      

    // If we are only showing today or if the date didn't change from the last item return the element for just the event
    if ((this.props.daysInAdvance == 0) || (!this._didDateChange(startDate, index))) {
      return itemElement;
    }

    return <div>
        <Text nowrap block className={styles.date}>{this._getDateText(startDate)}</Text>
        {itemElement}
      </div>;

    
  }

  private _getDateText = (date: Date): string => {
    var text: string;

    if (this._isDateToday(date)) {
      text = "Today";
    }

    else if (this._isDateTomorrow(date)) {
      text = "Tomorrow";
    }
    else if (this._isDateThisWeek(date)) {
      text = this._getDayOfWeek(date);
    }

    else {
      text = date.toLocaleDateString();
    }

    return text;
  }

  private _getDayOfWeek = (date: Date): string => {    
    switch (date.getDay()) {
      case 0:
        return "Sunday";
      case 1:
        return "Monday";
      case 2:
        return "Tuesday";
      case 3:
        return "Wednesday";
      case 4:
        return "Thursday";
      case 5:
        return "Friday";
      case 6:
        return "Saturday";      
    }
    return "ERROR";
  }

  private _didDateChange = (startDate: Date, index: number): boolean => {
    if (index == 0)
      return true;

    const previousItemDate: Date = new Date(this.state.meetings[index - 1].start.dateTime + 'Z');
    return !this._isDateSame(startDate, previousItemDate);    
  }

  private _getDateOnly = (date: Date | undefined): Date => {
    const dateOnly = (date == null || date == undefined) ? new Date() : new Date(date.getFullYear(), date.getMonth(), date.getDate());
    dateOnly.setHours(0);
    dateOnly.setMinutes(0);
    dateOnly.setSeconds(0);
    dateOnly.setMilliseconds(0);    
    return dateOnly;
  }

  private _isDateSame = (date1: Date, date2: Date): boolean => {
    return ((date1.getFullYear() == date2.getFullYear()) && (date1.getMonth() == date2.getMonth()) && (date1.getDate() == date2.getDate()));
  }

  private _isDateToday = (date: Date): boolean => {
    return this._isDateSame(date, new Date());    
  }  

  private _isDateTomorrow = (date: Date): boolean => {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    return this._isDateSame(date, tomorrow);
  }
  
  private _isDateThisWeek = (date: Date): boolean => {
    const nextWeek = new Date();
    nextWeek.setDate(nextWeek.getDate() + 7);
    return this._getDateOnly(date) < this._getDateOnly(nextWeek);
  }
  

  

  /**
     * Get user-friendly string that represents the duration of the meeting
     * < 1h: x minutes
     * >= 1h: 1 hour (y minutes)
     * all day: All day
     */
  private _getDuration = (meeting: IMeeting): string => {
    if (meeting.isAllDay) {
      return strings.AllDay;
    }

    const startDateTime: Date = new Date(meeting.start.dateTime);
    const endDateTime: Date = new Date(meeting.end.dateTime);
    // get duration in minutes
    const duration: number = Math.round((endDateTime as any) - (startDateTime as any)) / (1000 * 60);
    if (duration <= 0) {
      return '';
    }

    if (duration < 60) {
      return `${duration} ${strings.Minutes}`;
    }

    const hours: number = Math.floor(duration / 60);
    const minutes: number = Math.round(duration % 60);
    let durationString: string = `${hours} ${hours > 1 ? strings.Hours : strings.Hour}`;
    if (minutes > 0) {
      durationString += ` ${minutes} ${strings.Minutes}`;
    }

    return durationString;
  }

  private _getFilterEndDateString = (): string => {
    const date: Date = new Date();    
    date.setHours(23);
    date.setMinutes(59);
    date.setSeconds(0);
    date.setDate(date.getDate() + (this.props.daysInAdvance || 0));
    return date.toISOString();
  }

  /**
   * Forces re-render of the component
   */
  private _reRender = (): void => {
    // update the render date to force reloading data and re-rendering
    // the component
    this.setState({ renderedDateTime: new Date() });
  }

  /**
   * Sets interval so that the data in the component is refreshed on the
   * specified cycle
   */
  private _setInterval = (): void => {
    let { refreshInterval } = this.props;
    // set up safe default if the specified interval is not a number
    // or beyond the valid range
    if (isNaN(refreshInterval) || refreshInterval < 0 || refreshInterval > 60) {
      refreshInterval = 5;
    }
    // refresh the component every x minutes
    this._interval = window.setInterval(this._reRender, refreshInterval * 1000 * 60);
    this._reRender();
  }

  public componentDidMount(): void {
    this._setInterval();
  }

  private _onNewMeeting = (): void => {
    window.open("https://outlook.office.com/?path=/calendar/action/compose", "_blank");
  }

  public componentWillUnmount(): void {
    // remove the interval so that the data won't be reloaded
    clearInterval(this._interval);
  }

  public componentDidUpdate(prevProps: IMyCalendarProps, prevState: IMyCalendarState): void {
    // if the refresh interval has changed, clear the previous interval
    // and setup new one, which will also automatically re-render the component
    if (prevProps.refreshInterval !== this.props.refreshInterval) {
      clearInterval(this._interval);
      this._setInterval();
      return;
    }

    // reload data on new render interval
    if (prevState.renderedDateTime !== this.state.renderedDateTime) {
      this._loadCalendar();
    }
  }

  private _showEventDetails = (messageId: string): void => {
    this
      ._getEventDetails(messageId)
      .then((activeEvent: Event): void => {
        this.setState({
          isOpen: true,
          activeEvent: activeEvent
        });
      });
  }

  private _getEventDetails = (messageId: string): Promise<Event> => {
    return new Promise<Event>((resolve, reject) => {
      this._getTimeZone().then(timeZone => {
        this.props.graphClient
          // get the mailbox settings
          .api(`me/calendar/events/` + messageId)
          .version("v1.0")
          .header("Prefer", "outlook.timezone=" + '"' + timeZone + '"')
          .get((err: any, res: Event): void => {
            if (err) {
              console.log("error:" + err);
              return reject(err);
            }
            resolve(res);
          });
      });
    });
  }


  public render(): React.ReactElement<IMyCalendarProps> {

    const recipientStackTokens: IStackTokens = {
      childrenGap: 7
    };

    const eventDetailsStackTokens: IStackTokens = {
      childrenGap: 3
    };

    const messageDetailsCommandBarFarItems: ICommandBarItemProps[] = [
      {
        key: 'viewInToDo',
        text: 'View in ToDo',
        iconProps: { iconName: 'OpenInNewWindow' },
        href: this.state.activeEvent.webLink,
        target: '_blank'
      }
    ];

    return (
      <div className={styles.myCalendar}>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty} />
        {
          !this.state.loading &&
          <>
            <PrimaryButton iconProps={{ iconName: 'AddEvent' }} onClick={this._onNewMeeting} disabled={false} >
              {strings.NewMeeting}
            </PrimaryButton>
            <div className={styles.list}>
              <List items={this.state.meetings}
                onRenderCell={this._onRenderCell} className={styles.list} />
              <Link href='https://outlook.office.com/owa/?path=/calendar/view/Day' target='_blank'>{strings.ViewAll}</Link>
            </div>
          </>
        }
        {
          (true && this.state.isOpen) ?
            <Panel
              className={styles.eventDetails}
              isLightDismiss
              isOpen={this.state.isOpen}
              onDismiss={() => { this.setState({ isOpen: false }); }}
              type={PanelType.largeFixed}
              closeButtonAriaLabel="Close"
              headerText={this.state.activeEvent.subject}
            >
              <Stack tokens={recipientStackTokens}>
                <CommandBar
                  items={[]}
                  farItems={messageDetailsCommandBarFarItems}
                  className={styles.commandBar}
                  ariaLabel="Use left and right arrow keys to navigate between commands"
                />
                <Text>
                  {new Date(this.state.activeEvent.start.dateTime).toLocaleString()} - {new Date(this.state.activeEvent.end.dateTime).toLocaleTimeString()}
                </Text>
                {(this.state.activeEvent.body.contentType === "html") ?
                  <div dangerouslySetInnerHTML={{ __html: this.state.activeEvent.body.content }}></div> :
                  <div>{this.state.activeEvent.body.content}</div>
                }
              </Stack>
            </Panel> : null
        }
        {
          !this.state.loading &&
          this.state.error &&
          <>
            <span className={styles.error}>{this.state.error}</span> :
            <span className={styles.noMeetings}>{strings.NoMeetings}</span>
          </>
        }
      </div>
    );
  }
}
