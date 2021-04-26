// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as React from 'react';
import styles from './MyCalendar.module.scss';
import * as strings from 'MyCalendarWebPartStrings';
import { IMyCalendarProps } from './IMyCalendarProps';
import { IMyCalendarState } from './IMyCalendarState';
import { Providers } from '@microsoft/mgt';
import { Agenda, MgtTemplateProps } from '@microsoft/mgt-react';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { Link } from 'office-ui-fabric-react/lib/components/Link';
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
      Providers.globalProvider.graph
        // get the mailbox settings
        .api(`me/mailboxSettings`)
        .version("v1.0")
        .get((err: any, res: microsoftgraph.MailboxSettings): void => {
          if (err) {
            return reject(err);
          }

          resolve(res.timeZone);
        });
    });
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
    this
      ._getTimeZone()
      .then((_timeZone: string): void => {
        this.setState({
          timeZone: _timeZone,
          loading: false
        });
      });
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
      Providers.globalProvider.graph
        // get the mailbox settings
        .api(`me/calendar/events/` + messageId)
        .version("v1.0")
        .get((err: any, res: Event): void => {
          if (err) {
            console.log("error:" + err);
            return reject(err);
          }
          resolve(res);
        });
    });
  }

  private getDuration = (_event: Event): string => {
    if (_event.isAllDay) {
      return strings.AllDay;
    }

    const _startDateTime: Date = new Date(_event.start.dateTime);
    const _endDateTime: Date = new Date(_event.end.dateTime);
    // get duration in minutes
    const _duration: number = Math.round((_endDateTime as any) - (_startDateTime as any)) / (1000 * 60);
    if (_duration <= 0) {
      return '';
    }

    if (_duration < 60) {
      return `${_duration} ${strings.Minutes}`;
    }

    const _hours: number = Math.floor(_duration / 60);
    const _minutes: number = Math.round(_duration % 60);
    let durationString: string = `${_hours} ${_hours > 1 ? strings.Hours : strings.Hour}`;
    if (_minutes > 0) {
      durationString += ` ${_minutes} ${strings.Minutes}`;
    }

    return durationString;
  }

  public render(): React.ReactElement<IMyCalendarProps> {

    const EventInfo = (props: MgtTemplateProps) => {

      const event: Event | undefined = props.dataContext ? props.dataContext.event : undefined;

      if (!event) {
        return <div />;
      }

      const startTime: Date = new Date(event.start.dateTime);
      const minutes: number = startTime.getMinutes();

      return <div className={`${styles.meetingWrapper} ${event.showAs}`}>
        <Link className={styles.meeting} onClick={() => this._showEventDetails(event.id)} href="">
          <div className={styles.linkWrapper}>
            <div className={styles.start}>{`${startTime.getHours()}:${minutes < 10 ? '0' + minutes : minutes}`}</div>
            <div>
              <div className={styles.subject}>{event.subject}</div>
              <div className={styles.duration}>{this.getDuration(event)}</div>
              <div className={styles.location}>{event.location.displayName}</div>
            </div>
          </div>
        </Link>
      </div>;
    };

    const date: Date = new Date();
    const now: string = date.toISOString();
    // set the date to midnight today to load all upcoming meetings for today
    date.setUTCHours(23);
    date.setUTCMinutes(59);
    date.setUTCSeconds(0);
    date.setDate(date.getDate() + (this.props.daysInAdvance || 0));
    const midnight: string = date.toISOString();

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
              <Agenda
                preferredTimezone={this.state.timeZone}
                eventQuery={`me/calendar/calendarView?startDateTime=${now}&endDateTime=${midnight}`}
                showMax={this.props.numMeetings > 0 ? this.props.numMeetings : undefined} >
                <EventInfo template='event' />
              </Agenda>
            </div>
            <Link href='https://outlook.office.com/owa/?path=/calendar/view/Day' target='_blank'>{strings.ViewAll}</Link>
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
              </Stack>
              {(this.state.activeEvent.body.contentType === "html") ?
                <div dangerouslySetInnerHTML={{ __html: this.state.activeEvent.body.content }}></div> :
                <div>{this.state.activeEvent.body.content}</div>
              }
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
