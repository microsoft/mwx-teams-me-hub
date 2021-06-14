// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as React from 'react';
import styles from './MyCalendar.module.scss';
import * as strings from 'MyCalendarWebPartStrings';
import { HeaderDisplay, ClickAction } from '../enums';
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

  private _getTimeText(date: Date) {
    const hour: number = date.getHours();
    const adjHour: number = (hour > 12) ? hour - 12 : hour;
    const amPm: string = ((hour >= 12) && (hour != 24)) ? "PM" : "AM";
    const minutes: number = date.getMinutes();
    return `${adjHour}:${minutes < 10 ? '0' + minutes : minutes} ${amPm}`;
  }  


  /**
     * Render meeting item
     */
  private _onRenderCell = (item: IMeeting, index: number | undefined): JSX.Element => {
    const startDate: Date = new Date(item.start.dateTime + 'Z');
    const endDate: Date = new Date(item.end.dateTime + 'Z');     

    const itemElement: JSX.Element = 
      <div className={`${styles.meetingWrapper} ${item.showAs}`}>
        <Link className={styles.meeting} onClick={() => this._showEventDetails(item.id, item.webLink)} target='_blank' >
          <div className={styles.timeColumn}>
            <Text nowrap block className={styles.start}>{this._getTimeText(startDate)}</Text>
            <Text nowrap block className={styles.duration}>{this._getDurationText(item)}</Text>            
          </div>
          <div className={styles.detailsColumn}>            
            <Text nowrap block className={styles.subject}>{item.subject}</Text> 
            <Text nowrap block className={styles.location}>{item.location.displayName}</Text>             
          </div>          
        </Link>
      </div>;    

    // If we are only showing today or if the date didn't change from the last item return the element for just the event
    if ((this.props.daysInAdvance == 0) || (!this._didDateChange(startDate, index))) {
      return itemElement;
    }

    return (
      <>
        <Text nowrap block className={styles.date}>{this._getDateText(startDate)}</Text>
        {itemElement}
      </>);

    
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

  private _getDurationText = (meeting: IMeeting): string => {
    if (meeting.isAllDay) {
      return strings.AllDay;
    }

    const startDate: Date = new Date(meeting.start.dateTime + 'Z');
    const endDate: Date = new Date(meeting.end.dateTime + 'Z');        
    // get duration in minutes
    const duration: number = Math.round((endDate as any) - (startDate as any)) / (1000 * 60);
    var durationText: string;

    if (duration <= 0) {
      durationText = '';
    }
    else if (duration < 60) {
      durationText = `${duration} ${strings.Minutes}`;
    }
    else if (duration === 60)  {
      durationText = `1 ${strings.Hour}`;
    }
    else {
      durationText = `${Math.round((duration / 60) * 10)/10} ${strings.Hours}`;
    }
    
    return durationText;
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

  private _showEventDetails = (messageId: string, webLink: string): void => {
    if (this.props.clickAction === ClickAction.OpenInOutlook) {
      window.open(webLink, "_blank");
      return;
    }
    
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
        {
          this.props.headerDisplay != HeaderDisplay.None &&           
          <WebPartTitle 
            displayMode={this.props.displayMode}                  
            title={this.props.title}
            className={ (this.props.headerDisplay == HeaderDisplay.Standard) ? `${styles.webPartTitle} ${styles.webPartTitleStandard}` : styles.webPartTitle}
            updateProperty={this.props.updateProperty} //className={styles.title}         
            moreLink={ (this.props.showViewAll) ? <Link href="https://outlook.office.com/owa/" target="_blank">See all</Link> : null } 
          />
        }
        {
          !this.state.loading && this.props.showNew &&          
          <PrimaryButton iconProps={{ iconName: 'AddEvent' }} onClick={this._onNewMeeting} disabled={false} >
            {strings.NewMeeting}
          </PrimaryButton>
        }
        {
          !this.state.loading && 
          <div className={styles.list}>
            <List items={this.state.meetings}
              onRenderCell={this._onRenderCell} className={styles.list} />            
            {
              this.props.showViewAll && this.props.headerDisplay == HeaderDisplay.None && 
              <Link href='https://outlook.office.com/owa/?path=/calendar/view/Day' target='_blank' className={styles.viewAll}>{strings.ViewAll}</Link>
            }
          </div>          
        }
        {
          this.state.isOpen &&
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
          </Panel>
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
