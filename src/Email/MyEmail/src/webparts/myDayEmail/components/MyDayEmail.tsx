import * as React from 'react';
import styles from './MyDayEmail.module.scss';
import * as strings from 'MyDayEmailWebPartStrings';
import { IMyDayEmailProps, IMyDayEmailState, IMessage, IMessages, IMessageDetails, IEmailAddress } from '.';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from 'office-ui-fabric-react/lib/Stack';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { List } from 'office-ui-fabric-react/lib/components/List';
import { Link } from 'office-ui-fabric-react/lib/components/Link';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import { PrimaryButton, IIconProps } from 'office-ui-fabric-react';
import { FontIcon } from 'office-ui-fabric-react/lib/Icon';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { initializeIcons } from '@uifabric/icons';
import { mergeStyles, mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { GraphRequest } from '@microsoft/sp-http';

export class MyDayEmail extends React.Component<IMyDayEmailProps, IMyDayEmailState> {

  private _interval: number;

  private _messageIconClass: string = mergeStyles({
    fontSize: 16,
    height: 16,
    width: 16,
    margin: '0',
  });

  private _classNames = mergeStyleSets({
    attachment: [this._messageIconClass],
    highImportance: [{ color: 'red' }, this._messageIconClass],
    lowImportance: [this._messageIconClass],
    flagged: [{ color: 'red' }, this._messageIconClass],
    completed: [{ color: 'green' }, this._messageIconClass]
  });

  

  constructor(props: IMyDayEmailProps) {
    super(props);

    initializeIcons();

    this.state = {
      messages: [],
      loading: false,
      error: undefined,
      renderedDateTime: new Date(),
      filter: strings.AllPivot,
      isOpen: false,
      activeMessage: {} as IMessageDetails
    };
  }

  /**
   * Load recent messages for the current user
   */
  private _loadMessages(): void {
    if (!this.props.graphClient) {
      return;
    }

    // update state to indicate loading and remove any previously loaded
    // messages
    this.setState({
      error: null,
      loading: true,
      messages: []
    });   

    const request: GraphRequest = this.props.graphClient
      .api("me/mailFolders/Inbox/messages")
      .version("v1.0")
      .select("id,bodyPreview,receivedDateTime,from,subject,webLink,isRead,importance,flag,hasAttachments") //,meetingMessageType             
      .top(this.props.nrOfMessages || 5);

    // Graph API does not like ordering when we are viewing only flagged items.
    if ((this.state.filter != strings.FlaggedPivot) && (this.state.filter != strings.ImportantPivot)) {
        request.orderby("receivedDateTime desc");
    }

    if (this.state.filter == strings.UnreadPivot) {
      request.filter("isRead eq false");
    }
    else if (this.state.filter == strings.ImportantPivot) {
      request.filter("importance eq 'high'");
    }
    else if (this.state.filter == strings.FlaggedPivot) {
      request.filter("flag/flagStatus eq 'flagged'");
    }
      
    console.log(`email request: filter = ${this.state.filter} - url: ${request.buildFullUrl()}`);

    request.get((err: any, res: IMessages): void => {
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
          messages: res.value,
          loading: false
        });
      }
      else {
        // No messages found
        this.setState({
          loading: false
        });
      }
    });
    
  }  
  private _handlePivotChange = (item: PivotItem): void => {
    console.log(`Pivot Change: ${item.props.itemKey}`);

    this.setState({ 
      filter: item.props.itemKey 
    });

    this._reRender();
  }

  private _onNewEmail = (): void => {
    window.open("https://outlook.office.com/?path=/mail/action/compose", "_blank");
  }

  private _showMessageDetails = (messageId: string): void => {    
    const request: GraphRequest = this.props.graphClient
      .api(`me/mailFolders/Inbox/messages/${messageId}`)
      .version("v1.0")
      .select("id,bodyPreview,receivedDateTime,from,subject,webLink,isRead,importance,flag,hasAttachments,body,toRecipients,ccRecipients"); //,meetingMessageType   
      
      console.log(`email request: ${request.buildFullUrl()}`);

      request.get((err: any, res: IMessageDetails): void => {
        if (err) {
          // Something failed calling the MS Graph
          this.setState({
            error: err.message ? err.message : strings.Error,            
          });          
        }
        else if (res) {
          this.setState({
            isOpen: true,
            activeMessage: res          
          });

          console.log(`email content: ${this.state.activeMessage.body.content}`);
        }        
      });
  }

  private _reRender = (): void => {
    //this._loadMessages();

    // update the render date to force reloading data and re-rendering
    // the component    
    this.setState({ renderedDateTime: new Date() });
  }

  private _setInterval = (): void => {
    let { refreshInterval } = this.props;
    // set up safe default if the specified interval is not a number
    // or beyond the valid range
    if (isNaN(refreshInterval) || refreshInterval < 0 || refreshInterval > 60) {
      refreshInterval = 5;
    }
    // refresh the component every x minutes
    this._interval = setInterval(this._reRender, refreshInterval * 1000 * 60);
    this._reRender();
  }

 

  public componentDidMount(): void {
    // load data initially after the component has been instantiated
    this._loadMessages();
    this._setInterval();
  }

  public componentWillUnmount(): void {
    // remove the interval so that the data won't be reloaded
    clearInterval(this._interval);
  }

  public componentDidUpdate(prevProps: IMyDayEmailProps, prevState: IMyDayEmailState): void {

    if (prevProps.refreshInterval !== this.props.refreshInterval) {
      clearInterval(this._interval);
      this._setInterval();
      return;
    }

    // verify if the component should update. Helps avoid unnecessary re-renders
    // when the parent has changed but this component hasn't
    if ((prevProps.nrOfMessages !== this.props.nrOfMessages) || 
        (prevState.renderedDateTime !== this.state.renderedDateTime) || 
        (prevState.filter !== this.state.filter)) {
      this._loadMessages();
    }
  }

  public render(): React.ReactElement<IMyDayEmailProps> {
    
    const recipientStackTokens: IStackTokens = {
      childrenGap: 7    
    };
  
    const messageDetailsStackTokens: IStackTokens = {
      childrenGap: 3    
    };

    const messageDetailsCommandBarFarItems: ICommandBarItemProps[] = [      
      {
        key: 'viewInOutlook',
        text: 'View in Outlook',
        iconProps: { iconName: 'OpenInNewWindow' },
        href: this.state.activeMessage.webLink,
        target: '_blank'
      }
    ];

    return (
      <div className={styles.myDayEmail}>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty} className={styles.title} />
        {
          this.state.loading &&
          <Spinner label={strings.Loading} size={SpinnerSize.large} />
        }

        {
          this.state.messages &&
            this.state.messages.length > 0 ? (
              <div>                                          
              { ///TODO: Make displaying the button a web part property
                (true) ? 
                //don't render the button till we figure out how to go straight to new mail
                <PrimaryButton iconProps={{iconName: 'NewMail'}} onClick={this._onNewEmail} disabled={false} >
                  New message
                </PrimaryButton> : null
              }
              <Pivot
                initialSelectedKey={strings.AllPivot} 
                onLinkClick={this._handlePivotChange}
                selectedKey={this.state.filter}
                headersOnly={true}>
                <PivotItem headerText={strings.AllPivot} itemKey={strings.AllPivot} />
                <PivotItem headerText={strings.UnreadPivot} itemKey={strings.UnreadPivot} />                
              </Pivot>              
              <List items={this.state.messages}
                onRenderCell={this._onRenderEmailCell} className={styles.list} />
              <Link href='https://outlook.office.com/owa/' target='_blank' className={styles.viewAll}>{strings.ViewAll}</Link>
            </div>
            ) : (
              !this.state.loading && (
                this.state.error ?
                  <span className={styles.error}>{this.state.error}</span> :
                  <span className={styles.noMessages}>{strings.NoMessages}</span>
              )
            )
        }
        {
          (true && this.state.isOpen) ? 
        <Panel
          className={styles.messageDetails}
          isLightDismiss          
          isOpen={this.state.isOpen}
          onDismiss={() => { this.setState( { isOpen: false }); } }
          type={PanelType.largeFixed}          
          closeButtonAriaLabel="Close"
          headerText={this.state.activeMessage.subject}
        >   
          <Stack tokens={messageDetailsStackTokens}>            
            <CommandBar
              items={[]}
              farItems={messageDetailsCommandBarFarItems}
              className={styles.commandBar}
              
              ariaLabel="Use left and right arrow keys to navigate between commands"
            />
            <Text>
              {this.state.activeMessage.from.emailAddress.name || this.state.activeMessage.from.emailAddress.address}
            </Text>          
            <Text>
              {new Date(this.state.activeMessage.receivedDateTime).toLocaleDateString()} 
              {new Date(this.state.activeMessage.receivedDateTime).toLocaleTimeString()}
            </Text>          
            <Stack horizontal disableShrink className={styles.recipients} tokens={recipientStackTokens}>
                <Text className={styles.prompt}>To: </Text>
                {this.state.activeMessage.toRecipients.map((recipient, index) =>(
                  <Text className={styles.recipient}>
                    {(recipient.emailAddress.name || recipient.emailAddress.address)}
                    {(index !== this.state.activeMessage.toRecipients.length -1) ? "; " : ""}
                  </Text>      
                ))}             
            </Stack>
            {  (this.state.activeMessage.ccRecipients != null && this.state.activeMessage.ccRecipients.length > 0) ?
            <Stack horizontal disableShrink className={styles.recipients} tokens={recipientStackTokens}>
                <Text className={styles.prompt}>Cc: </Text>
                {this.state.activeMessage.ccRecipients.map((recipient, index) =>(
                  <Text className={styles.recipient}>
                    {(recipient.emailAddress.name || recipient.emailAddress.address)}
                    {(index !== this.state.activeMessage.ccRecipients.length -1) ? "; " : ""}
                  </Text>      
                ))}             
            </Stack> : null
            }                  
            {  (this.state.activeMessage.body.contentType === "html") ?           
            <div dangerouslySetInnerHTML={{ __html: this.state.activeMessage.body.content }}></div> : 
            <div>{this.state.activeMessage.body.content}</div>
            }        
          </Stack>    
        </Panel> : null
      }
      </div>
    );
  }
  
   /**
   * Render message item
   */
  private _onRenderEmailCell = (item: IMessage, index: number | undefined): JSX.Element => {
    var messageStyle:string = styles.message;

    if (item.isRead) {
      messageStyle = styles.message + " " + styles.isRead;
    }

    return (
      <Link className={messageStyle} onClick={ () => this._showMessageDetails(item.id) }>
         <div className={styles.from}>
          { ((item.from.emailAddress.name || item.from.emailAddress.address).length > 35) ?
          (item.from.emailAddress.name || item.from.emailAddress.address).substr(0, 32) + '...' :
          (item.from.emailAddress.name || item.from.emailAddress.address)
          }
        </div>
        <div className={styles.icons}>
          {(item.hasAttachments) ? <FontIcon iconName="Attach" className={this._classNames.attachment} />: null}          
          {(item.importance == "high") ? <FontIcon iconName="Important" className={this._classNames.highImportance} /> : null}
          {(item.importance == "low") ? <FontIcon iconName="SortDown" className={this._classNames.lowImportance} /> : null}
          {(item.flag.flagStatus == "flagged") ? <FontIcon iconName="Flag" className={this._classNames.flagged} /> : null }
          {(item.flag.flagStatus == "complete") ? <FontIcon iconName="CheckMark" className={this._classNames.completed} />: null }          
        </div>
        <div className={styles.subject}>{item.subject}</div>
        <div className={styles.date}>{(new Date(item.receivedDateTime).toLocaleDateString())}</div>
        <div className={styles.preview}>{item.bodyPreview}</div>
      </Link>
    );
  } 
}
