// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as React from 'react';
import styles from './MyCalendar.module.scss';
import * as strings from 'MyCalendarWebPartStrings';
import { IMeeting, IMeetings } from './IMeeting';
import { IMyCalendarState } from './IMyCalendarState';
import { IMyCalendarProps } from './IMyCalendarProps';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { Link } from 'office-ui-fabric-react/lib/components/Link';
import { List } from 'office-ui-fabric-react/lib/components/List';
import { Event } from '@microsoft/microsoft-graph-types';
import { PrimaryButton, IIconProps } from 'office-ui-fabric-react';
import { initializeIcons } from '@uifabric/icons';
import { mergeStyles, mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from 'office-ui-fabric-react/lib/Stack';
import { Text } from 'office-ui-fabric-react/lib/Text';


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
  private _loadMeetings(): void {
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

    const date: Date = new Date();
    const now: string = date.toISOString();
    // set the date to midnight today to load all upcoming meetings for today
    date.setUTCHours(23);
    date.setUTCMinutes(59);
    date.setUTCSeconds(0);
    date.setDate(date.getDate() + (this.props.daysInAdvance || 0));
    const midnight: string = date.toISOString();

    this._getTimeZone().then(timeZone => {
      this.props.graphClient
        // get all upcoming meetings for the rest of the day today
        .api(`me/calendar/calendarView?startDateTime=${now}&endDateTime=${midnight}`)
        .version("v1.0")
        .select('subject,start,end,showAs,webLink,location,isAllDay')
        .top(this.props.numMeetings > 0 ? this.props.numMeetings : 100)
        .header("Prefer", "outlook.timezone=" + '"' + timeZone + '"')
        // sort ascending by start time
        .orderby("start/dateTime")
        .get((err: any, res: IMeetings): void => {
          if (err) {
            // Something failed calling the MS Graph
            this.setState({
              error: err.message ? err.message : strings.Error,
              loading: false
            });
            return;
          }

          // Check if a response was retrieved
          if (res && res.value && res.value.length > 0) {
            this.setState({
              meetings: res.value,
              loading: false
            });
          }
          else {
            // No meetings found
            this.setState({
              loading: false
            });
          }
        });
    });
  }


  /**
     * Render meeting item
     */
  private _onRenderCell = (item: IMeeting, index: number | undefined): JSX.Element => {
    const startTime: Date = new Date(item.start.dateTime);
    const minutes: number = startTime.getMinutes();

    return <div className={`${styles.meetingWrapper} ${item.showAs}`}>
      <Link className={styles.meeting} onClick={() => this._showEventDetails(item.id)} target='_blank' >
        <div className={styles.start}>{`${startTime.getHours()}:${minutes < 10 ? '0' + minutes : minutes}`}</div>
        <div className={styles.subject}>{item.subject}</div>
        <div className={styles.duration}>{this._getDuration(item)}</div>
        <div className={styles.location}>{item.location.displayName}</div>
      </Link>
    </div>;
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
      this._loadMeetings();
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
